import os
import sys
import io
import unittest
from pathlib import Path
import pandas as pd
import openpyxl

PROJECT_ROOT = Path(__file__).resolve().parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from core.models import (
    CATEGORY_GAP,
    CATEGORY_PHANTOM,
    CATEGORY_ACTIVE,
    CATEGORY_HEALTHY,
    CATEGORY_UNTRACKED,
    ReconciliationSummary,
    ReconciliationResult,
)
from core.normalizer import (
    normalize_ocpp_id,
    parse_issue_age,
    parse_datetime,
)
from core.loader import (
    detect_header_row,
    load_dataset,
)
from core.reconciler import (
    reconcile_incident_records,
)
from core.exporter import (
    generate_reconciliation_excel,
    export_reconciliation_csv,
)

SAMPLE_FIR_PATH = Path("D:/Comperator/DUMP/FIR DUMP.xlsx")
SAMPLE_OMS_PATH = Path("D:/Comperator/DUMP/OMS DUMP.xlsx")


class TestNormalizationAndParsing(unittest.TestCase):

    def test_normalize_ocpp_id_series(self):
        raw_series = pd.Series([
            " RJ-0267 ",
            "jk_0004",
            "UP 0183",
            "MH-0244",
            "nan",
            "None",
            "-",
            None,
        ])
        expected = [
            "RJ0267",
            "JK0004",
            "UP0183",
            "MH0244",
            "",
            "",
            "",
            "",
        ]
        result = normalize_ocpp_id(raw_series, strict=False)
        self.assertEqual(list(result), expected)

    def test_normalize_ocpp_id_single_value(self):
        self.assertEqual(normalize_ocpp_id(" RJ-0267 "), "RJ0267")
        self.assertEqual(normalize_ocpp_id("jk_0004"), "JK0004")
        self.assertEqual(normalize_ocpp_id(None), "")
        self.assertEqual(normalize_ocpp_id("-"), "")
        self.assertEqual(normalize_ocpp_id("RJ-0267", strict=True), "RJ-0267")

    def test_parse_issue_age(self):
        self.assertAlmostEqual(parse_issue_age("1 Days 23 Hours"), 1.958, places=3)
        self.assertAlmostEqual(parse_issue_age("0 Days 04 Hours"), 0.167, places=3)
        self.assertAlmostEqual(parse_issue_age("4 Days 12 Hours"), 4.5, places=3)
        self.assertAlmostEqual(parse_issue_age("4 Days"), 4.0, places=3)
        self.assertAlmostEqual(parse_issue_age("1 Day"), 1.0, places=3)
        self.assertAlmostEqual(parse_issue_age("12 Hours"), 0.5, places=3)
        self.assertAlmostEqual(parse_issue_age("6 Hours"), 0.25, places=3)
        self.assertAlmostEqual(parse_issue_age(3.5), 3.5, places=3)
        self.assertAlmostEqual(parse_issue_age("2.75"), 2.75, places=3)
        self.assertEqual(parse_issue_age(None), 0.0)
        self.assertEqual(parse_issue_age("nan"), 0.0)
        self.assertEqual(parse_issue_age("-"), 0.0)

    def test_detect_header_row_real_dumps(self):
        self.assertTrue(SAMPLE_FIR_PATH.exists(), f"Missing {SAMPLE_FIR_PATH}")
        self.assertTrue(SAMPLE_OMS_PATH.exists(), f"Missing {SAMPLE_OMS_PATH}")

        fir_header = detect_header_row(SAMPLE_FIR_PATH, filename="FIR DUMP.xlsx")
        self.assertEqual(fir_header, 2)

        oms_header = detect_header_row(SAMPLE_OMS_PATH, filename="OMS DUMP.xlsx")
        self.assertEqual(oms_header, 1)


class TestEndToEndReconciliation(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        cls.fir_df = load_dataset(SAMPLE_FIR_PATH, filename="FIR DUMP.xlsx")
        cls.oms_df = load_dataset(SAMPLE_OMS_PATH, filename="OMS DUMP.xlsx")
        cls.result = reconcile_incident_records(cls.fir_df, cls.oms_df)

    def test_total_fir_counts(self):
        summary = self.result.summary
        self.assertEqual(summary.total_fir, 54)
        self.assertEqual(summary.in_maintenance, 38)
        self.assertEqual(summary.available_cleared, 16)
        self.assertEqual(summary.in_maintenance + summary.available_cleared, summary.total_fir)

    def test_reconciliation_exact_categorization(self):
        summary = self.result.summary
        self.assertEqual(summary.ticket_gaps, 5)
        self.assertEqual(summary.phantom_resolutions, 13)
        self.assertEqual(summary.active_verified, 20)
        self.assertAlmostEqual(summary.coverage_gap_rate, 47.37, delta=0.1)

    def test_untracked_oms_open_tickets(self):
        summary = self.result.summary
        self.assertEqual(summary.untracked_oms_open, 478)

    def test_partitioned_dataframe_lengths_and_schemas(self):
        res = self.result
        self.assertEqual(len(res.reconciled_df), 54)
        self.assertEqual(len(res.gaps_df), 5)
        self.assertEqual(len(res.phantoms_df), 13)
        self.assertEqual(len(res.active_df), 20)
        self.assertEqual(len(res.healthy_df), 16)
        self.assertEqual(len(res.untracked_oms_df), 478)

        required_cols = [
            "Audit_Category",
            "Risk_Level",
            "Prescribed_Action",
            "Latest_Ticket_ID",
            "Latest_Ticket_Status",
            "Parsed_Issue_Age_Days",
        ]
        for col in required_cols:
            self.assertIn(col, res.reconciled_df.columns)

    def test_known_ticket_generation_gap_chargers(self):
        gaps_ocpp = set(self.result.gaps_df["_norm_ocpp"])
        expected_gaps = {"RJ0267", "SMC043", "JK0002", "TS0032", "SMC037"}
        self.assertEqual(gaps_ocpp, expected_gaps)


class TestMultiagentRemovalAudit(unittest.TestCase):

    def test_legacy_files_do_not_exist(self):
        forbidden_files = [
            PROJECT_ROOT / "orchestrator.py",
            PROJECT_ROOT / "multiagent_framework.py",
            PROJECT_ROOT / "test_multiagent_system.py",
        ]
        for f in forbidden_files:
            self.assertFalse(f.exists(), f"Forbidden file still exists: {f}")

    def test_no_forbidden_references_in_codebase(self):
        forbidden_tokens = [
            "multiagent_framework",
            "WarRoomMultiAgentSystem",
            "ReconciliationOrchestrator",
            "from orchestrator",
            "import orchestrator",
        ]

        py_files = [
            p for p in PROJECT_ROOT.rglob("*.py")
            if "tests" not in p.parts and ".git" not in p.parts
        ]

        for py_path in py_files:
            content = py_path.read_text(encoding="utf-8")
            for token in forbidden_tokens:
                self.assertNotIn(
                    token,
                    content,
                    f"Forbidden token '{token}' found in {py_path}",
                )


class TestExports(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        cls.fir_df = load_dataset(SAMPLE_FIR_PATH, filename="FIR DUMP.xlsx")
        cls.oms_df = load_dataset(SAMPLE_OMS_PATH, filename="OMS DUMP.xlsx")
        cls.result = reconcile_incident_records(cls.fir_df, cls.oms_df)

    def test_generate_reconciliation_excel(self):
        excel_bytes = generate_reconciliation_excel(self.result)
        self.assertIsInstance(excel_bytes, bytes)
        self.assertGreater(len(excel_bytes), 20000)

        wb = openpyxl.load_workbook(io.BytesIO(excel_bytes))
        sheet_names = wb.sheetnames
        expected_sheets = [
            "Executive Summary",
            "Ticket Generation Gaps",
            "Phantom Resolutions",
            "Active Verified",
            "Master Reconciled",
            "OMS Untracked Open",
        ]
        for s in expected_sheets:
            self.assertIn(s, sheet_names, f"Expected sheet '{s}' missing in generated Excel workbook")

    def test_export_reconciliation_csv(self):
        csv_text = export_reconciliation_csv(self.result.gaps_df)
        self.assertIsInstance(csv_text, str)
        lines = csv_text.strip().splitlines()
        self.assertEqual(len(lines), 6)


class TestStreamlitSmoke(unittest.TestCase):

    def test_app_syntax_and_compilation(self):
        app_path = PROJECT_ROOT / "app.py"
        self.assertTrue(app_path.exists(), "app.py does not exist")
        code = app_path.read_text(encoding="utf-8")
        compiled = compile(code, str(app_path), "exec")
        self.assertIsNotNone(compiled)

    def test_streamlit_config_exists(self):
        config_path = PROJECT_ROOT / ".streamlit" / "config.toml"
        self.assertTrue(config_path.exists(), ".streamlit/config.toml is missing")


class TestEdgeCaseRobustness(unittest.TestCase):

    def test_empty_oms_dataframe_input(self):
        fir_df = pd.DataFrame({"OCPP ID": ["C1"], "Status": ["Faulted"]})
        oms_df = pd.DataFrame()
        result = reconcile_incident_records(fir_df, oms_df)
        self.assertEqual(result.summary.total_fir, 1)
        self.assertEqual(result.summary.ticket_gaps, 1)
        self.assertEqual(result.summary.untracked_oms_open, 0)
        self.assertEqual(len(result.untracked_oms_df), 0)
        self.assertIn("Audit_Category", result.untracked_oms_df.columns)

    def test_mixed_timezone_strings_in_oms_created_at(self):
        fir_df = pd.DataFrame({"OCPP ID": ["C1"], "Status": ["Faulted"]})
        oms_df = pd.DataFrame({
            "Charger": ["C1", "C1"],
            "Status": ["OPEN", "OPEN"],
            "Created At": ["2024-01-01 10:00:00+05:30", "2024-01-01 10:00:00Z"],
        })
        result = reconcile_incident_records(fir_df, oms_df)
        self.assertEqual(result.summary.active_verified, 1)
        self.assertEqual(result.summary.ticket_gaps, 0)

    def test_search_filtering_resilience_missing_columns(self):
        df_all = pd.DataFrame({
            "Some_Other_Col": ["Alpha", "Beta"],
            "Audit_Category": [CATEGORY_GAP, CATEGORY_ACTIVE],
        })
        search_query = "alpha"
        q = search_query.lower()
        search_filter = pd.Series(False, index=df_all.index)
        if "OCPP ID" in df_all.columns:
            search_filter |= df_all["OCPP ID"].astype(str).str.lower().str.contains(q)
        if "Station Name" in df_all.columns:
            search_filter |= df_all["Station Name"].astype(str).str.lower().str.contains(q)
        filtered = df_all[search_filter]
        self.assertEqual(len(filtered), 0)


if __name__ == "__main__":
    unittest.main(verbosity=2)
