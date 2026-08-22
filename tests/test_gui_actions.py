from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace

from src.gui import SodSatApplication


class GuiActionStateTests(unittest.TestCase):
    def _file(self, path: Path) -> Path:
        path.write_text("ok", encoding="utf-8")
        return path

    def test_missing_system_reconciliation_enables_no_result_actions(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            base = Path(temporary)
            result = SimpleNamespace(status="BLOQUEADO_SEM_CONCILIACAO", workbook_path=None, sod_analysis_path=None, diagnostics_path=self._file(base / "diagnostico.csv"), email_required=False, email_draft_path=self._file(base / "minuta.md"))
            states = SodSatApplication._result_action_states(result)
            self.assertEqual(states, {"open": "disabled", "download": "disabled", "diagnostics": "disabled", "email": "disabled", "sod": "disabled"})

    def test_generated_result_only_enables_email_when_profile_divergence_exists(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            base = Path(temporary)
            result = SimpleNamespace(status="GERADO_SEM_PENDENCIAS", workbook_path=self._file(base / "resultado.xlsx"), sod_analysis_path=self._file(base / "sod.xlsx"), diagnostics_path=self._file(base / "diagnostico.csv"), email_required=False, email_draft_path=self._file(base / "minuta.md"))
            states = SodSatApplication._result_action_states(result)
            self.assertEqual(states["open"], "normal")
            self.assertEqual(states["download"], "normal")
            self.assertEqual(states["sod"], "normal")
            self.assertEqual(states["email"], "disabled")
            result.email_required = True
            self.assertEqual(SodSatApplication._result_action_states(result)["email"], "normal")

    def test_profile_divergence_enables_only_treatment_actions(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            base = Path(temporary)
            result = SimpleNamespace(status="BLOQUEADO_POR_DIVERGENCIA_DE_PERFIS", workbook_path=None, sod_analysis_path=self._file(base / "sod.xlsx"), diagnostics_path=self._file(base / "diagnostico.csv"), email_required=True, email_draft_path=self._file(base / "minuta.md"))
            self.assertEqual(SodSatApplication._result_action_states(result), {"open": "disabled", "download": "disabled", "diagnostics": "normal", "email": "normal", "sod": "normal"})


if __name__ == "__main__":
    unittest.main()

