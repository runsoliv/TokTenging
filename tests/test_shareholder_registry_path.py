from pathlib import Path
import tempfile
import unittest

from tok import shareholder_registry


class RegistryWorkbookPathTests(unittest.TestCase):
    def setUp(self):
        self.original_workbook_path = shareholder_registry.registry_workbook_path
        self.original_training_data_dir = shareholder_registry._training_data_dir
        self.original_resource_path = shareholder_registry.resource_path
        shareholder_registry.registry_workbook_path = None

    def tearDown(self):
        shareholder_registry.registry_workbook_path = self.original_workbook_path
        shareholder_registry._training_data_dir = self.original_training_data_dir
        shareholder_registry.resource_path = self.original_resource_path

    def test_default_workbook_is_found_beside_imports_folder(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            training_data = Path(temp_dir) / "trainingData"
            (training_data / "imports").mkdir(parents=True)
            data_folder = training_data / "data folder"
            data_folder.mkdir()
            workbook = data_folder / "data codex.xlsx"
            workbook.touch()

            shareholder_registry._training_data_dir = lambda: str(training_data)
            shareholder_registry.resource_path = lambda relative: str(
                Path(temp_dir) / "bundle" / relative
            )

            resolved = shareholder_registry._resolve_default_workbook()

        self.assertEqual(resolved, str(workbook.resolve()))

    def test_bundled_training_data_workbook_is_the_next_fallback(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            missing_training_data = Path(temp_dir) / "configured" / "trainingData"
            bundle_root = Path(temp_dir) / "bundle"
            workbook = (
                bundle_root
                / "tok"
                / "trainingData"
                / "data folder"
                / "data codex.xlsx"
            )
            workbook.parent.mkdir(parents=True)
            workbook.touch()

            shareholder_registry._training_data_dir = lambda: str(missing_training_data)
            shareholder_registry.resource_path = lambda relative: str(bundle_root / relative)

            resolved = shareholder_registry._resolve_default_workbook()

        self.assertEqual(resolved, str(workbook.resolve()))


if __name__ == "__main__":
    unittest.main()
