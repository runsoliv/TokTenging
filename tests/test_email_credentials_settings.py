import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import tok.TokTenging as app


class EmailCredentialSettingsTests(unittest.TestCase):
    def test_yahoo_credentials_survive_settings_reload(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            settings_path = Path(temp_dir) / "settings.json"
            training_dir = Path(temp_dir) / "training"
            saved_settings = {
                "yahoo_email": "account@example.com",
                "yahoo_app_password": "example-app-password",
                "training_data_dir": str(training_dir),
                "auto_code_mode": "ml_hybrid",
            }
            with patch.object(app, "SETTINGS_PATH", str(settings_path)):
                with patch.object(app, "settings", saved_settings):
                    app.save_settings()
                reloaded = app.load_settings()

            self.assertEqual(reloaded["yahoo_email"], "account@example.com")
            self.assertEqual(reloaded["yahoo_app_password"], "example-app-password")
            self.assertEqual(reloaded["auto_code_mode"], "ml_hybrid")

    def test_invalid_auto_code_mode_fails_closed_to_legacy(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            settings_path = Path(temp_dir) / "settings.json"
            training_dir = Path(temp_dir) / "training"
            settings_path.write_text(
                '{"auto_code_mode": "unknown", "training_data_dir": "'
                + str(training_dir).replace("\\", "\\\\")
                + '"}',
                encoding="utf-8",
            )
            with patch.object(app, "SETTINGS_PATH", str(settings_path)):
                reloaded = app.load_settings()

        self.assertEqual(reloaded["auto_code_mode"], "legacy")


if __name__ == "__main__":
    unittest.main()
