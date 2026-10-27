"""Small, dependency-light classical models for Tok Tenging autocoding.

The production application cannot assume that scikit-learn is installed.  This
module therefore implements two interpretable CPU-friendly baselines with
NumPy, which is already a pandas dependency:

* a character/word n-gram TF-IDF nearest-centroid classifier; and
* a character/word n-gram complement naive-Bayes classifier.

Both use deterministic feature hashing.  They are deliberately modest models
for the roughly ten-thousand reviewed examples available to this project, and
they expose calibrated top candidates rather than forcing every prediction.
"""

from __future__ import annotations

from collections import Counter, defaultdict
from dataclasses import asdict, dataclass
from typing import Any, Iterable, Mapping, Sequence
import math
import re
import zlib

import numpy as np

try:
    from .ml_features import FEATURE_VERSION, build_feature_document
except ImportError:  # pragma: no cover - direct script compatibility
    from ml_features import FEATURE_VERSION, build_feature_document


MODEL_SCHEMA_VERSION = "tok-local-ngram-v1"
_WORD_RE = re.compile(r"[a-z0-9_]+")


@dataclass(frozen=True)
class NgramModelConfig:
    name: str
    algorithm: str
    dimensions: int = 32768
    word_ngram_max: int = 2
    char_ngram_min: int = 3
    char_ngram_max: int = 5
    use_character_ngrams: bool = True
    sublinear_tf: bool = True
    alpha: float = 0.35
    prior_strength: float = 0.02
    maximum_document_characters: int = 1600
    similarity_neighbors: int = 9
    similarity_candidates: int = 250
    similarity_max_document_frequency: float = 0.25
    similarity_power: float = 4.0

    def validate(self) -> None:
        if self.algorithm not in {"centroid", "complement_nb", "similarity"}:
            raise ValueError(f"unsupported model algorithm: {self.algorithm}")
        if self.dimensions < 1024:
            raise ValueError("feature hash dimension is too small")
        if self.word_ngram_max not in {1, 2, 3}:
            raise ValueError("word_ngram_max must be between one and three")
        if not 1 <= self.char_ngram_min <= self.char_ngram_max <= 7:
            raise ValueError("invalid character n-gram range")
        if self.alpha <= 0:
            raise ValueError("alpha must be positive")
        if self.similarity_neighbors < 1 or self.similarity_candidates < self.similarity_neighbors:
            raise ValueError("invalid similarity neighbor/candidate limits")
        if not 0.0 < self.similarity_max_document_frequency <= 1.0:
            raise ValueError("similarity document-frequency limit must be in (0, 1]")
        if self.similarity_power <= 0:
            raise ValueError("similarity power must be positive")


DEFAULT_CANDIDATE_CONFIGS = (
    NgramModelConfig(
        name="tfidf_char_word_centroid",
        algorithm="centroid",
        prior_strength=0.015,
    ),
    NgramModelConfig(
        name="char_word_complement_nb",
        algorithm="complement_nb",
        char_ngram_max=4,
        sublinear_tf=False,
        alpha=0.45,
        prior_strength=0.0,
    ),
    NgramModelConfig(
        name="char_word_similarity_knn",
        algorithm="similarity",
        char_ngram_max=4,
        sublinear_tf=False,
        prior_strength=0.0,
    ),
)


@dataclass(frozen=True)
class ModelPrediction:
    predicted: str
    probability: float
    margin: float
    candidates: tuple[tuple[str, float], ...]
    class_support: int

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


def _stable_feature_index(value: str, dimensions: int) -> int:
    # crc32 is deterministic between processes and materially faster than a
    # cryptographic hash for the millions of short n-grams used in evaluation.
    return zlib.crc32(value.encode("utf-8")) % dimensions


def _feature_names(document: str, config: NgramModelConfig) -> Iterable[str]:
    document = str(document or "")[: config.maximum_document_characters]
    words = _WORD_RE.findall(document.lower())
    for size in range(1, config.word_ngram_max + 1):
        for position in range(0, len(words) - size + 1):
            yield f"w{size}:" + " ".join(words[position:position + size])

    if not config.use_character_ngrams or not words:
        return
    # Boundaries preserve short merchant tokens.  Replacing separators with a
    # single space makes punctuation/reference formatting differences benign.
    character_text = f"^{ ' '.join(words) }$"
    for size in range(config.char_ngram_min, config.char_ngram_max + 1):
        for position in range(0, len(character_text) - size + 1):
            yield f"c{size}:" + character_text[position:position + size]


def hashed_document(document: str, config: NgramModelConfig) -> tuple[np.ndarray, np.ndarray]:
    counts: Counter[int] = Counter(
        _stable_feature_index(feature, config.dimensions)
        for feature in _feature_names(document, config)
    )
    if not counts:
        return np.asarray([], dtype=np.int32), np.asarray([], dtype=np.float32)
    indices = np.fromiter(sorted(counts), dtype=np.int32)
    values = np.fromiter((counts[int(index)] for index in indices), dtype=np.float32)
    if config.sublinear_tf:
        values = 1.0 + np.log(values)
    return indices, values


def feature_documents(records: Any) -> list[str]:
    if hasattr(records, "to_dict"):
        records = records.to_dict("records")
    return [build_feature_document(record) for record in records]


class LocalNgramClassifier:
    """Pickle-safe local n-gram classifier with explicit probability scaling."""

    def __init__(self, config: NgramModelConfig):
        config.validate()
        self.config = config
        self.schema_version = MODEL_SCHEMA_VERSION
        self.feature_version = FEATURE_VERSION
        self.classes_: np.ndarray = np.asarray([], dtype=str)
        self.class_support_: np.ndarray = np.asarray([], dtype=np.int64)
        self.idf_: np.ndarray | None = None
        self.weights_: np.ndarray | None = None
        self.log_prior_: np.ndarray | None = None
        self.temperature_: float = 1.0
        self.training_vectors_: list[np.ndarray] = []
        self.training_labels_: np.ndarray = np.asarray([], dtype=np.int32)
        self.training_norms_: np.ndarray = np.asarray([], dtype=np.float32)
        self.inverted_index_: dict[int, np.ndarray] = {}

    def ready(self) -> bool:
        return (
            self.schema_version == MODEL_SCHEMA_VERSION
            and self.feature_version == FEATURE_VERSION
            and len(self.classes_) >= 2
            and (
                self.weights_ is not None
                or (
                    self.config.algorithm == "similarity"
                    and bool(self.inverted_index_)
                    and len(self.training_labels_) == int(self.class_support_.sum())
                )
            )
        )

    def fit(
        self,
        documents: Sequence[str],
        labels: Sequence[Any],
        *,
        temperature: float | None = None,
    ) -> "LocalNgramClassifier":
        if len(documents) != len(labels):
            raise ValueError("documents and labels must have equal length")
        if len(documents) < 2:
            raise ValueError("at least two training documents are required")
        labels_array = np.asarray([str(value) for value in labels], dtype=str)
        classes, encoded = np.unique(labels_array, return_inverse=True)
        if len(classes) < 2:
            raise ValueError("at least two target classes are required")

        vectors = [hashed_document(document, self.config) for document in documents]
        self.classes_ = classes
        self.class_support_ = np.bincount(encoded, minlength=len(classes)).astype(np.int64)
        priors = self.class_support_.astype(np.float64) / len(documents)
        self.log_prior_ = np.log(np.maximum(priors, 1e-12)).astype(np.float32)
        if self.config.algorithm == "centroid":
            self._fit_centroid(vectors, encoded)
        elif self.config.algorithm == "complement_nb":
            self._fit_complement_nb(vectors, encoded)
        else:
            self._fit_similarity(vectors, encoded)
        if temperature is not None:
            self.temperature_ = self._validated_temperature(temperature)
        return self

    def fit_records(
        self,
        records: Any,
        labels: Sequence[Any] | None = None,
        *,
        label_column: str = "model_target",
        temperature: float | None = None,
    ) -> "LocalNgramClassifier":
        if labels is None:
            if hasattr(records, "__getitem__") and hasattr(records, "columns"):
                labels = records[label_column].astype(str).tolist()
            else:
                raise ValueError("labels are required for non-dataframe records")
        return self.fit(feature_documents(records), labels, temperature=temperature)

    def _fit_centroid(
        self,
        vectors: Sequence[tuple[np.ndarray, np.ndarray]],
        encoded: np.ndarray,
    ) -> None:
        document_frequency = np.zeros(self.config.dimensions, dtype=np.uint32)
        for indices, _values in vectors:
            if len(indices):
                document_frequency[indices] += 1
        self.idf_ = (
            np.log((1.0 + len(vectors)) / (1.0 + document_frequency.astype(np.float64))) + 1.0
        ).astype(np.float32)
        centroids = np.zeros((len(self.classes_), self.config.dimensions), dtype=np.float32)
        for class_index, (indices, values) in zip(encoded, vectors):
            if not len(indices):
                continue
            weighted = values * self.idf_[indices]
            norm = float(np.linalg.norm(weighted))
            if norm:
                centroids[int(class_index), indices] += weighted / norm
        centroids /= np.maximum(self.class_support_[:, None], 1)
        norms = np.linalg.norm(centroids, axis=1)
        nonzero = norms > 0
        centroids[nonzero] /= norms[nonzero, None]
        self.weights_ = centroids

    def _fit_complement_nb(
        self,
        vectors: Sequence[tuple[np.ndarray, np.ndarray]],
        encoded: np.ndarray,
    ) -> None:
        class_counts = np.zeros((len(self.classes_), self.config.dimensions), dtype=np.float32)
        for class_index, (indices, values) in zip(encoded, vectors):
            if len(indices):
                class_counts[int(class_index), indices] += values
        all_counts = class_counts.sum(axis=0, dtype=np.float64)
        complement = all_counts[None, :] - class_counts.astype(np.float64)
        smoothed = complement + float(self.config.alpha)
        denominators = smoothed.sum(axis=1, keepdims=True)
        # Complement NB predicts the class with the smallest complement
        # likelihood.  Negating normalized log probabilities turns that into
        # the standard largest-score interface used by this class.
        weights = -np.log(smoothed / np.maximum(denominators, 1e-12))
        row_norms = np.abs(weights).sum(axis=1, keepdims=True)
        weights /= np.maximum(row_norms, 1e-12)
        self.weights_ = weights.astype(np.float32)
        self.idf_ = None

    def _fit_similarity(
        self,
        vectors: Sequence[tuple[np.ndarray, np.ndarray]],
        encoded: np.ndarray,
    ) -> None:
        document_frequency = np.zeros(self.config.dimensions, dtype=np.uint32)
        for indices, _values in vectors:
            if len(indices):
                document_frequency[indices] += 1
        self.idf_ = (
            np.log((1.0 + len(vectors)) / (1.0 + document_frequency.astype(np.float64))) + 1.0
        ).astype(np.float32)
        maximum_frequency = max(
            2,
            int(math.ceil(len(vectors) * self.config.similarity_max_document_frequency)),
        )
        retained = (document_frequency > 0) & (document_frequency <= maximum_frequency)
        postings: dict[int, list[int]] = {}
        training_vectors = []
        training_norms = np.zeros(len(vectors), dtype=np.float32)
        for document_index, (indices, _values) in enumerate(vectors):
            filtered = indices[retained[indices]] if len(indices) else indices
            filtered = np.unique(filtered).astype(np.int32, copy=False)
            training_vectors.append(filtered)
            if len(filtered):
                weights = self.idf_[filtered]
                training_norms[document_index] = float(np.sqrt(np.sum(weights * weights)))
                for feature_index in filtered:
                    postings.setdefault(int(feature_index), []).append(document_index)
        self.training_vectors_ = training_vectors
        self.training_labels_ = encoded.astype(np.int32, copy=True)
        self.training_norms_ = training_norms
        self.inverted_index_ = {
            feature_index: np.asarray(document_indices, dtype=np.int32)
            for feature_index, document_indices in postings.items()
        }
        self.weights_ = None

    def decision_function(self, documents: Sequence[str]) -> np.ndarray:
        if not self.ready():
            raise RuntimeError("model has not been fitted")
        if self.config.algorithm == "similarity":
            probabilities = self._similarity_probabilities(documents)
            return np.log(np.maximum(probabilities, 1e-12)).astype(np.float32)
        scores = np.zeros((len(documents), len(self.classes_)), dtype=np.float32)
        for row_index, document in enumerate(documents):
            indices, values = hashed_document(document, self.config)
            if len(indices):
                if self.config.algorithm == "centroid":
                    weighted = values * self.idf_[indices]
                    norm = float(np.linalg.norm(weighted))
                    if norm:
                        weighted = weighted / norm
                    scores[row_index] = self.weights_[:, indices] @ weighted
                else:
                    scores[row_index] = self.weights_[:, indices] @ values
            if self.log_prior_ is not None and self.config.prior_strength:
                scores[row_index] += self.config.prior_strength * self.log_prior_
        return scores

    def predict_proba(self, documents: Sequence[str]) -> np.ndarray:
        if self.config.algorithm == "similarity":
            return self._similarity_probabilities(documents)
        scores = self.decision_function(documents)
        scaled = scores / self._validated_temperature(self.temperature_)
        scaled -= scaled.max(axis=1, keepdims=True)
        probabilities = np.exp(np.clip(scaled, -80.0, 0.0))
        denominators = probabilities.sum(axis=1, keepdims=True)
        return probabilities / np.maximum(denominators, 1e-12)

    def _similarity_probabilities(self, documents: Sequence[str]) -> np.ndarray:
        if not self.ready():
            raise RuntimeError("model has not been fitted")
        class_priors = self.class_support_.astype(np.float64)
        class_priors /= max(1.0, float(class_priors.sum()))
        output = np.zeros((len(documents), len(self.classes_)), dtype=np.float64)
        for row_index, document in enumerate(documents):
            indices, _values = hashed_document(document, self.config)
            filtered = np.asarray(
                sorted({int(index) for index in indices if int(index) in self.inverted_index_}),
                dtype=np.int32,
            )
            if not len(filtered):
                output[row_index] = class_priors
                continue
            feature_weights = self.idf_[filtered].astype(np.float64) ** 2
            query_norm = math.sqrt(float(feature_weights.sum()))
            candidate_scores: dict[int, float] = defaultdict(float)
            for feature_index, feature_weight in zip(filtered, feature_weights):
                for document_index in self.inverted_index_[int(feature_index)]:
                    candidate_scores[int(document_index)] += float(feature_weight)
            if not candidate_scores or not query_norm:
                output[row_index] = class_priors
                continue
            strongest = sorted(
                candidate_scores.items(),
                key=lambda item: (item[1], -item[0]),
                reverse=True,
            )[: self.config.similarity_candidates]
            similarities = []
            for document_index, dot_product in strongest:
                denominator = query_norm * float(self.training_norms_[document_index])
                if denominator <= 0:
                    continue
                similarity = max(0.0, min(1.0, dot_product / denominator))
                similarities.append((similarity, document_index))
            if not similarities:
                output[row_index] = class_priors
                continue
            similarities.sort(key=lambda item: (item[0], -item[1]), reverse=True)
            maximum_similarity = similarities[0][0]
            cutoff = max(0.25, maximum_similarity * 0.75)
            neighbors = [
                item for item in similarities
                if item[0] >= cutoff
            ][: self.config.similarity_neighbors]
            votes = np.zeros(len(self.classes_), dtype=np.float64)
            for similarity, document_index in neighbors:
                votes[int(self.training_labels_[document_index])] += similarity ** self.config.similarity_power
            vote_total = float(votes.sum())
            if vote_total <= 0:
                output[row_index] = class_priors
                continue
            vote_distribution = votes / vote_total
            evidence = max(0.0, min(1.0, (maximum_similarity - 0.35) / 0.60))
            probabilities = (evidence * vote_distribution) + ((1.0 - evidence) * class_priors)
            output[row_index] = probabilities / max(float(probabilities.sum()), 1e-12)
        return output.astype(np.float32)

    def predict(self, documents: Sequence[str], *, top_k: int = 3) -> list[ModelPrediction]:
        probabilities = self.predict_proba(documents)
        top_k = max(1, min(int(top_k), len(self.classes_)))
        predictions = []
        for row in probabilities:
            ranked = np.argsort(-row, kind="stable")[:top_k]
            best = int(ranked[0])
            second_probability = float(row[int(ranked[1])]) if len(ranked) > 1 else 0.0
            candidates = tuple(
                (str(self.classes_[int(index)]), float(row[int(index)]))
                for index in ranked
            )
            predictions.append(ModelPrediction(
                predicted=str(self.classes_[best]),
                probability=float(row[best]),
                margin=float(row[best]) - second_probability,
                candidates=candidates,
                class_support=int(self.class_support_[best]),
            ))
        return predictions

    def predict_records(self, records: Any, *, top_k: int = 3) -> list[ModelPrediction]:
        return self.predict(feature_documents(records), top_k=top_k)

    def calibrate_temperature(
        self,
        documents: Sequence[str],
        labels: Sequence[Any],
    ) -> float:
        if self.config.algorithm == "similarity":
            self.temperature_ = 1.0
            return self.temperature_
        if len(documents) != len(labels):
            raise ValueError("calibration documents and labels must have equal length")
        if not documents:
            return self.temperature_
        class_index = {str(value): index for index, value in enumerate(self.classes_)}
        known_positions = [position for position, label in enumerate(labels) if str(label) in class_index]
        if not known_positions:
            return self.temperature_
        known_documents = [documents[position] for position in known_positions]
        targets = np.asarray([class_index[str(labels[position])] for position in known_positions], dtype=np.int64)
        scores = self.decision_function(known_documents).astype(np.float64)
        # The normalized complement-NB score scale is much smaller than a
        # cosine-centroid scale.  A logarithmic grid supports both without
        # pretending their raw margins are comparable probabilities.
        grid = np.geomspace(1e-5, 100.0, num=57)
        best_temperature = self.temperature_
        best_loss = math.inf
        row_indices = np.arange(len(targets))
        for temperature in grid:
            scaled = scores / temperature
            scaled -= scaled.max(axis=1, keepdims=True)
            log_denominator = np.log(np.exp(np.clip(scaled, -80.0, 0.0)).sum(axis=1))
            loss = float(np.mean(log_denominator - scaled[row_indices, targets]))
            if loss < best_loss:
                best_loss = loss
                best_temperature = float(temperature)
        self.temperature_ = self._validated_temperature(best_temperature)
        return self.temperature_

    @staticmethod
    def _validated_temperature(value: Any) -> float:
        try:
            temperature = float(value)
        except (TypeError, ValueError) as exc:
            raise ValueError("invalid probability temperature") from exc
        if not math.isfinite(temperature) or temperature <= 0:
            raise ValueError("probability temperature must be positive and finite")
        return max(1e-6, min(100.0, temperature))

    def metadata(self) -> dict[str, Any]:
        return {
            "schema_version": self.schema_version,
            "feature_version": self.feature_version,
            "config": asdict(self.config),
            "class_count": int(len(self.classes_)),
            "training_rows": int(self.class_support_.sum()) if len(self.class_support_) else 0,
            "temperature": float(self.temperature_),
        }


class LocalModelEnsemble:
    """Conservative agreement ensemble over independently evaluated models."""

    def __init__(self, models: Sequence[LocalNgramClassifier], name: str = "unanimous_ensemble"):
        if len(models) < 2:
            raise ValueError("an ensemble needs at least two models")
        self.models = list(models)
        self.name = str(name)
        self.schema_version = MODEL_SCHEMA_VERSION
        self.feature_version = FEATURE_VERSION
        self.classes_ = self.models[0].classes_.copy()
        self.class_support_ = self.models[0].class_support_.copy()

    def ready(self) -> bool:
        if not all(model.ready() for model in self.models):
            return False
        expected = self.classes_.tolist()
        return all(model.classes_.tolist() == expected for model in self.models)

    def predict(self, documents: Sequence[str], *, top_k: int = 3) -> list[ModelPrediction]:
        if not self.ready():
            raise RuntimeError("ensemble models are not ready or class-compatible")
        component_predictions = [model.predict(documents, top_k=top_k) for model in self.models]
        output = []
        for row_index in range(len(documents)):
            row_predictions = [predictions[row_index] for predictions in component_predictions]
            votes = Counter(prediction.predicted for prediction in row_predictions)
            predicted, vote_count = votes.most_common(1)[0]
            if vote_count == len(row_predictions):
                probability = 0.999
                margin = 0.90
            elif vote_count >= math.ceil(len(row_predictions) / 2):
                probability = 0.85
                margin = 0.40
            else:
                strongest = max(row_predictions, key=lambda item: item.probability)
                predicted = strongest.predicted
                probability = min(0.49, strongest.probability)
                margin = min(0.10, strongest.margin)
            candidate_scores: Counter[str] = Counter()
            for prediction in row_predictions:
                for code, candidate_probability in prediction.candidates:
                    candidate_scores[str(code)] += float(candidate_probability)
            ranked = candidate_scores.most_common(max(1, min(top_k, len(self.classes_))))
            total = sum(score for _code, score in ranked) or 1.0
            candidates = [(predicted, probability)]
            candidates.extend(
                (code, min(probability, score / total))
                for code, score in ranked
                if code != predicted
            )
            candidates = tuple(candidates[:top_k])
            class_index = int(np.where(self.classes_ == predicted)[0][0])
            output.append(ModelPrediction(
                predicted=predicted,
                probability=float(probability),
                margin=float(margin),
                candidates=candidates,
                class_support=int(self.class_support_[class_index]),
            ))
        return output

    def predict_records(self, records: Any, *, top_k: int = 3) -> list[ModelPrediction]:
        return self.predict(feature_documents(records), top_k=top_k)

    def metadata(self) -> dict[str, Any]:
        return {
            "schema_version": self.schema_version,
            "feature_version": self.feature_version,
            "name": self.name,
            "class_count": int(len(self.classes_)),
            "training_rows": int(self.class_support_.sum()),
            "components": [model.metadata() for model in self.models],
        }


def candidate_models() -> list[LocalNgramClassifier]:
    return [LocalNgramClassifier(config) for config in DEFAULT_CANDIDATE_CONFIGS]


__all__ = [
    "DEFAULT_CANDIDATE_CONFIGS",
    "MODEL_SCHEMA_VERSION",
    "LocalNgramClassifier",
    "LocalModelEnsemble",
    "ModelPrediction",
    "NgramModelConfig",
    "candidate_models",
    "feature_documents",
    "hashed_document",
]
