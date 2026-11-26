"""Tier 4 E2E Tests: Real-World Operational Scenarios (6 multi-feature journeys)."""

from __future__ import annotations

import csv
from datetime import date
import io
import unittest

import pandas as pd

from tests_e2e.e2e_harness import (
    AppSession,
    VIEW_OMS_DASHBOARD,
    VIEW_OMS_FIR_COMPARE,
    assert_reference_converted_df,
    load_xlsx_workbook_bytes,
    parse_csv_bytes,
)
from tests_e2e.fixtures import (
    ALT_TODAY,
    DATE_ONLY_RESOLVED_COL,
    DISK_FIR_PATH,
    DISK_OMS_PATH,
    EXPECTED_HOLD_DELTAS_2026_12_01,
    RAW_RESOLVED_COL,
    REFERENCE_42_COLUMNS_DATE_ONLY,
    REFERENCE_TODAY,
    build_reference_9_rows_df,
    to_csv_bytes,
    to_xls_xml_bytes,
    to_xlsx_bytes,
)


def _apply_filter(session: AppSession, option_value: str) -> None:
    session.select_dashboard_filter(option_value)


class TestTier4Scenarios(unittest.TestCase):
    def test_t4_01_reference_9_row_nov1_25_oms_end_to_end(self) -> None:
        """T4-01: Full 9-row Nov1-25-OMS -> Nov1-25-OMS_date_only UI, exports, and round-trip."""
        session = AppSession().start()
        self.assertEqual(session.get_nav_radio().value, VIEW_OMS_DASHBOARD)

        raw_df = build_reference_9_rows_df("iso")
        session.upload_to(
            session.find_dashboard_uploader(),
            "Nov1-25-OMS.xlsx",
            to_xlsx_bytes(raw_df),
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        session.set_date(session.find_today_date_input(), REFERENCE_TODAY)
        session.assert_no_exceptions()

        self.assertRegex(session.extract_kpi_text(), r"\b9\b")
        self.assertGreaterEqual(len(session.get_charts()), 1)

        csv_dl = session.find_download(suffix=".csv")
        xlsx_dl = session.find_download(suffix=".xlsx")

        csv_text = csv_dl["bytes"].decode("utf-8")
        self.assertNotIn(".0", csv_text)
        csv_df = parse_csv_bytes(csv_dl["bytes"])
        assert_reference_converted_df(csv_df)

        wb = load_xlsx_workbook_bytes(xlsx_dl["bytes"])
        ws = wb.active
        headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
        self.assertEqual(headers, list(REFERENCE_42_COLUMNS_DATE_ONLY))
        self.assertEqual(ws.max_row, 10)
        app_col = headers.index("Approval Time taken (Days)") + 1
        tid_col = headers.index("Ticket ID") + 1
        xlsx_app = {ws.cell(r, tid_col).value: ws.cell(r, app_col).value for r in range(2, ws.max_row + 1)}
        self.assertEqual(xlsx_app["EX-2026-32559"], -18)
        self.assertIsNone(xlsx_app["EX-2026-49803"])

        # Round-trip re-upload of exported CSV
        s2 = AppSession().start()
        s2.upload_to(s2.find_dashboard_uploader(), "Nov1-25-OMS_date_only.csv", csv_dl["bytes"], "text/csv")
        s2.set_date(s2.find_today_date_input(), REFERENCE_TODAY)
        s2.assert_no_exceptions()
        assert_reference_converted_df(parse_csv_bytes(s2.find_download(suffix=".csv")["bytes"]))

    def test_t4_02_full_disk_fir_oms_reconciliation_workflow(self) -> None:
        """T4-02: Full D:\\Comperator\\DUMP\\FIR DUMP.xlsx & OMS DUMP.xlsx reconciliation journey."""
        self.assertTrue(DISK_FIR_PATH.exists())
        self.assertTrue(DISK_OMS_PATH.exists())

        session = AppSession().start()
        session.switch_view(VIEW_OMS_FIR_COMPARE)
        session.click_button("Sample Files")
        self.assertIn("54 chargers found", session.extract_all_text())
        self.assertIn("2605 tickets found", session.extract_all_text())
        session.click_button("Compare Both Files")
        session.assert_no_exceptions()

        kpi = session.extract_kpi_text()
        for expected in ("54", "38", "5", "13", "20", "47.4%"):
            self.assertIn(expected, kpi)
        self.assertEqual(
            session.get_tab_labels(),
            [
                "Missing Tickets (5)",
                "Wrongly Closed Tickets (13)",
                "Being Worked On (20)",
                "Charts & Trends",
                "Extra Tickets in OMS (478)",
                "All Chargers List (54)",
            ],
        )

        session.select_pill_containing("Wrongly Closed")
        session.set_multiselect(session.find_multiselect("Region / City"), ["West"])
        session.set_text(session.find_search_input(), "Ador")
        session.assert_no_exceptions()

        wb = load_xlsx_workbook_bytes(session.find_download(suffix=".xlsx")["bytes"])
        self.assertIn("Phantom Resolutions", wb.sheetnames)
        wrongly_csv = parse_csv_bytes(session.find_download(filename_contains="wrongly_closed")["bytes"])
        self.assertGreaterEqual(len(wrongly_csv), 1)

        session.click_button("Reset Filters")
        self.assertIn("All Chargers List (54)", session.get_tab_labels())
        session.click_button("Clear All")
        session.assert_no_exceptions()
        self.assertIn("EV Charger Complaint & Status Checker", session.extract_all_text())

    def test_t4_03_full_2605_row_disk_oms_dump_in_oms_dashboard(self) -> None:
        """T4-03: Full 2,605-row D:\\Comperator\\DUMP\\OMS DUMP.xlsx in OMS Dashboard."""
        self.assertTrue(DISK_OMS_PATH.exists())
        session = AppSession().start()
        session.upload_to(
            session.find_dashboard_uploader(),
            "OMS DUMP.xlsx",
            DISK_OMS_PATH.read_bytes(),
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        session.assert_no_exceptions()

        kpi = session.extract_kpi_text()
        self.assertTrue("2,605" in kpi or "2605" in kpi)
        self.assertGreaterEqual(len(session.get_charts()), 1)

        csv_df = parse_csv_bytes(session.find_download(suffix=".csv")["bytes"])
        self.assertEqual(len(csv_df), 2605)
        ex_4574 = csv_df.loc[csv_df["Ticket ID"] == "EX-2026-4574"].iloc[0]
        self.assertEqual(str(ex_4574["Created At"]), "2026-09-30")
        self.assertEqual(str(ex_4574["Updated At"]), "2026-11-19")
        self.assertEqual(str(ex_4574["Closed At"]), "2026-11-19")

        wb = load_xlsx_workbook_bytes(session.find_download(suffix=".xlsx")["bytes"])
        self.assertEqual(wb.active.max_row, 2606)

        _apply_filter(session, "West")
        _apply_filter(session, "Critical")
        session.assert_no_exceptions()
        expected_ids = set(
            csv_df.loc[
                (csv_df["Zone"].astype(str) == "West") & (csv_df["Priority"].astype(str) == "Critical"),
                "Ticket ID",
            ].astype(str)
        )
        filtered_dfs = [d for d in session.get_dataframes() if "Ticket ID" in d.columns]
        narrowed = [d for d in filtered_dfs if 0 < len(d) < 2605]

        def _matches_west_critical(d: pd.DataFrame) -> bool:
            tids = set(d["Ticket ID"].astype(str))
            if not tids or not tids.issubset(expected_ids):
                return False
            if "Zone" in d.columns and set(d["Zone"].astype(str)) != {"West"}:
                return False
            if "Priority" in d.columns and set(d["Priority"].astype(str)) != {"Critical"}:
                return False
            return True

        self.assertTrue(
            any(_matches_west_critical(d) for d in narrowed),
            f"Expected narrowed Ticket ID table matching West+Critical; got shapes={[d.shape for d in filtered_dfs]}",
        )

    def test_t4_04_multi_shift_operations_triage_and_cross_view_handoff(self) -> None:
        """T4-04: Morning/afternoon shift date handoff on OMS Dashboard + Compare audit."""
        session = AppSession().start()
        raw_df = build_reference_9_rows_df("iso")
        session.upload_to(session.find_dashboard_uploader(), "Nov1-25-OMS.csv", to_csv_bytes(raw_df), "text/csv")

        # Morning shift: today = 2026-11-25, filter On Hold + Critical
        session.set_date(session.find_today_date_input(), REFERENCE_TODAY)
        _apply_filter(session, "On Hold")
        _apply_filter(session, "Critical")
        wb_morning = load_xlsx_workbook_bytes(session.find_download(suffix=".xlsx")["bytes"])
        ws_m = wb_morning.active
        hdrs_m = [ws_m.cell(1, c).value for c in range(1, ws_m.max_column + 1)]
        hold_m = {
            ws_m.cell(r, hdrs_m.index("Ticket ID") + 1).value: ws_m.cell(r, hdrs_m.index("Hold total time(Days)") + 1).value
            for r in range(2, ws_m.max_row + 1)
        }
        self.assertEqual(hold_m["EX-2026-46397"], 0)

        # Afternoon shift: today = 2026-11-26 (+1 day)
        afternoon = date(2026, 11, 26)
        session.set_date(session.find_today_date_input(), afternoon)
        afternoon_csv = parse_csv_bytes(session.find_download(suffix=".csv")["bytes"]).set_index("Ticket ID")
        self.assertEqual(int(afternoon_csv.loc["EX-2026-46397", "Hold total time(Days)"]), 1)

        # Switch to OMS-FIR Compare and run full audit
        session.switch_view(VIEW_OMS_FIR_COMPARE)
        session.click_button("Sample Files")
        session.click_button("Compare Both Files")
        audit_xlsx = session.find_download(suffix=".xlsx")
        self.assertIn("Executive Summary", load_xlsx_workbook_bytes(audit_xlsx["bytes"]).sheetnames)

        # Switch back to OMS Dashboard and verify afternoon state preserved
        session.switch_view(VIEW_OMS_DASHBOARD)
        self.assertEqual(session.find_today_date_input().value, afternoon)
        preserved_csv = parse_csv_bytes(session.find_download(suffix=".csv")["bytes"]).set_index("Ticket ID")
        self.assertEqual(int(preserved_csv.loc["EX-2026-46397", "Hold total time(Days)"]), 1)

    def test_t4_05_multi_format_batch_ingestion_and_reupload_isolation(self) -> None:
        """T4-05: Identical 9-row fixture uploaded as BOM .csv, .xlsx, and .xls yields identical exports."""
        raw_df = build_reference_9_rows_df("iso")
        session = AppSession().start()

        # 1. UTF-8-BOM CSV
        session.upload_to(
            session.find_dashboard_uploader(),
            "batch1.csv",
            to_csv_bytes(raw_df, bom=True),
            "text/csv",
        )
        session.set_date(session.find_today_date_input(), ALT_TODAY)
        csv_from_csv = session.find_download(suffix=".csv")["bytes"]

        # 2. XLSX replacement in same session
        session.upload_to(
            session.find_dashboard_uploader(),
            "batch2.xlsx",
            to_xlsx_bytes(raw_df),
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        session.assert_no_exceptions()
        csv_from_xlsx = session.find_download(suffix=".csv")["bytes"]

        # 3. XLS (SpreadsheetML XML) replacement in same session
        session.upload_to(
            session.find_dashboard_uploader(),
            "batch3.xls",
            to_xls_xml_bytes(raw_df),
            "application/vnd.ms-excel",
        )
        session.assert_no_exceptions()
        csv_from_xls = session.find_download(suffix=".csv")["bytes"]

        df_1 = parse_csv_bytes(csv_from_csv)
        df_2 = parse_csv_bytes(csv_from_xlsx)
        df_3 = parse_csv_bytes(csv_from_xls)
        assert_reference_converted_df(df_1, expected_hold_map=EXPECTED_HOLD_DELTAS_2026_12_01)
        pd.testing.assert_frame_equal(df_1, df_2)
        pd.testing.assert_frame_equal(df_1, df_3)

    def test_t4_06_dirty_production_dump_resilience_and_recovery(self) -> None:
        """T4-06: Pre-header banners, padded headers, mixed dates, sentinels, formulas, and special search."""
        raw_df_serial = build_reference_9_rows_df("mixed")
        iso_df = build_reference_9_rows_df("iso")
        raw_df_serial["Created At"] = iso_df["Created At"]
        raw_df_serial.loc[raw_df_serial.index[0], "Assigned At"] = "N/A"
        raw_df_serial.loc[raw_df_serial.index[0], "Hold Status Time"] = "-"
        raw_df_serial.loc[raw_df_serial.index[5], "Acknowledgement time taken(Days)"] = "=U7-S7"
        raw_df_serial.loc[raw_df_serial.index[0], "Station"] = "[O&M] Bay-42"
        dirty_xlsx = to_xlsx_bytes(
            raw_df_serial,
            excel_serial_cols=("Created At",),
            pre_header_rows=2,
            header_overrides={RAW_RESOLVED_COL: f"  {RAW_RESOLVED_COL}  "},
        )

        session = AppSession().start()
        session.upload_to(
            session.find_dashboard_uploader(),
            "dirty_oms_dump.xlsx",
            dirty_xlsx,
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        session.set_date(session.find_today_date_input(), REFERENCE_TODAY)
        session.assert_no_exceptions()

        session.set_text(session.find_search_input(), "[O&M]")
        session.assert_no_exceptions()
        matched = [d for d in session.get_dataframes() if "Ticket ID" in d.columns and len(d) == 1]
        self.assertTrue(matched)

        out_df = parse_csv_bytes(session.find_download(suffix=".csv")["bytes"])
        assert_reference_converted_df(out_df)


if __name__ == "__main__":
    unittest.main()
