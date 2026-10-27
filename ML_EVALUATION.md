# Tok Tenging ML autocoder evaluation

Evaluation date: 2026-08-31
Production gate: **FAILED — no ML model was activated**

## Reproduce

```powershell
.venv\Scripts\python.exe -m tok.ml_training evaluate --training-data-dir tok\trainingData --legacy-training-dir tok\trainingCoded --folds 3 --output artifacts\ml\evaluation_report.json
```

To evaluate and activate a newly trained model only when every production gate
passes:

```powershell
.venv\Scripts\python.exe -m tok.ml_training rebuild --training-data-dir tok\trainingData --legacy-training-dir tok\trainingCoded --folds 3 --output artifacts\ml\evaluation_report.json --artifact-dir tok\trainingData\ml_artifacts
```

With the current data this command reports `activated: false` and leaves the
active artifact untouched, because the precision gate does not pass.

The generated JSON report contains per-account, per-bank, grouped-split,
calibration, coverage, and confusion details. `artifacts/` is intentionally
ignored because generated financial-data/model outputs must not be committed.

## Verified input

- 77 manifest-designated import sessions and 77 readable coded workbooks.
- 10,266 coded rows inspected; the historical audit definition still finds
  8,927 compatible rows.
- 9,734 rows passed stricter direction, target, and bank-side validation before
  deduplication.
- 22 duplicate rows and both rows in one conflicting-label identity were
  removed, leaving 9,710 usable rows.
- 139 `islandsbanki_special` rows were retained for audit but excluded from the
  production-model scope, leaving 9,571 evaluation rows.
- 127 literal target codes; 119 model targets after the configured 4635/4640
  vehicle-expense family is collapsed.
- 8,062 raw links were identity/compression validated. Unsafe, ambiguous, and
  coded-only rows remain usable with raw-derived features missing.
- Class imbalance is severe: 18 singleton target classes, 43 classes with fewer
  than five examples, and 57 with fewer than ten.

## Unseen-session comparison

| System | Accuracy | Macro F1 | Weighted F1 | Top-2 | Top-3 |
|---|---:|---:|---:|---:|---:|
| Deterministic only | 7.52% | 2.05% | 7.98% | — | — |
| Exact verified history only | 13.74% | 7.41% | 20.64% | — | — |
| Current complete autocoder | 50.49% | 22.28% | 50.83% | — | — |
| Complement NB | 51.08% | 18.66% | 47.09% | 64.11% | 71.34% |
| Similarity model | 46.57% | 23.13% | 45.08% | 57.92% | 63.58% |
| Centroid model | 48.49% | 22.56% | 46.63% | 63.50% | 73.01% |
| Unanimous ensemble | 55.41% | 24.86% | 53.08% | 68.60% | 74.57% |

The ensemble improved raw top-1 accuracy over the current autocoder, but its
apparently high-confidence predictions were not trustworthy enough to automate.
The held-out confidence policy accepted 753 unresolved rows (25.87% coverage)
at only 52.59% precision, far below the 95% production target. Account 6220 was
also unsafe: ensemble precision was 48.46%, versus 64.08% for the current
autocoder.

The resulting policy has `production_enabled=false`. In all three required gate
schemes (unseen session, unseen client, and unseen merchant), the production
hybrid therefore accepts zero ML rows and returns the legacy result unchanged.

## Production behavior

The persisted settings switch offers `Hefðbundin` and `ML-aðstoð`. The legacy
autocoder always runs first. ML may only replace unresolved/review rows after a
checksum-validated, compatible artifact and its confidence policy both pass the
production gate. Manual/MN, deterministic/AD, user-selected 4635/4640 behavior,
and consistent exact history are protected. Missing, corrupt, incompatible, or
failed ML inference returns the complete legacy frame without partial changes.

Because this evaluation failed the gate, there is deliberately no active model
artifact. Selecting `ML-aðstoð` currently uses the legacy result and the settings
page reports that fallback in Icelandic.

## Most valuable next labels

Prioritize independently reviewed examples for unseen clients/merchants and the
largest observed confusions: 4082/4600 versus the vehicle-expense family, 4082
versus 4061, 9410/4510 versus 4061, and 9325 versus 9410. Add explicit negative
examples around generic invoice text and collection-cost/6220 transactions, and
raise support for the 57 classes currently below ten examples. Re-run the same
grouped command; activate ML only after accepted-row precision reaches the gate
without reducing the legacy result.
