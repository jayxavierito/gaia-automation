import unittest
from pathlib import Path

from gaia_ui_automation import TrialConfig, validate_trial_config


class UiAutomationConfigTests(unittest.TestCase):
    def setUp(self):
        self.root = Path.cwd()
        self.candidate = self.root / "test_ui_candidate.xlsx"
        self.source_review = self.root / "test_ui_candidate_review.csv"
        self.executable = self.root / "test_GaiaCloud.exe"
        for path in (self.candidate, self.source_review, self.executable):
            path.touch()

    def tearDown(self):
        for path in (self.candidate, self.source_review, self.executable):
            path.unlink(missing_ok=True)

    def make_config(self, project_name: str) -> TrialConfig:
        candidate = self.candidate
        source_review = self.source_review
        executable = self.executable
        for path in (candidate, source_review, executable):
            path.touch()
        return TrialConfig(
            candidate=candidate,
            source_review=source_review,
            project_name=project_name,
            folder_name="テスト",
            output_dir=self.root / "test_ui_output",
            gaia_exe=executable,
            cache_root=self.root / "test_ui_cache",
            timeout_seconds=60,
            confirm_start=True,
            export_review=False,
        )

    def test_accepts_explicit_test_project(self):
        config = self.make_config("GAIA自動化テスト_20260722_example")

        validate_trial_config(config)

    def test_rejects_live_project_name(self):
        config = self.make_config("live_project")

        with self.assertRaisesRegex(ValueError, "GAIA自動化テスト_"):
            validate_trial_config(config)


if __name__ == "__main__":
    unittest.main()
