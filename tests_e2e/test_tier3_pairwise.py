"""Tier 3 E2E Tests: Cross-Feature Pairwise Combinations (15 tests)."""

from __future__ import annotations

from datetime import date
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
    DELTA_COLS_8,
    DISK_OMS_PATH,
    EXPECTED_HOLD_DELTAS_2026_12_01,
    REFERENCE_TODAY,
    build_reference_9_rows_df,
    corrupt_binary_bytes,
    to_csv_bytes,
    to_xls_xml_bytes,
    to_xlsx_bytes,
)


def _apply_filter(session: AppSession, option_value: str) -> None:
    session.select_dashboard_filter(option_value)


class TestTier3Pairwise(unittest.TestCase):
    def test_t3_01_mixed_date_csv_with_formula_hints_full_pipeline(self) -> None:
        session = AppSession().start()
        raw_df = build_reference_9_rows_df("mixed")
        session.upload_to(session.find_dashboard_uploader(), "Nov1-25-mixed.csv", to_csv_bytes(raw_df), "text/csv")
        session.set_date(session.find_today_date_input(), REFERENCE_TODAY)
        session.assert_no_exceptions()
        out_df = parse_csv_bytes(session.find_download(suffix=".csv")["bytes"])
        assert_reference_converted_df(out_df)

    def test_t3_02_xlsx_upload_today_switch_2026_11_25_and_2026_12_01_exports(self) -> None:
        session = AppSession().start()
        raw_df = build_reference_9_rows_df("iso")
        session.upload_to(
            session.find_dashboard_uploader(),
            "Nov1-25-OMS.xlsx",
            to_xlsx_bytes(raw_df),
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        session.set_date(session.find_today_date_input(), REFERENCE_TODAY)
        assert_reference_converted_df(parse_csv_bytes(session.find_download(suffix=".csv")["bytes"]))

        session.set_date(session.find_today_date_input(), ALT_TODAY)
        csv_df = parse_csv_bytes(session.find_download(suffix=".csv")["bytes"])
        assert_reference_converted_df(csv_df, expected_hold_map=EXPECTED_HOLD_DELTAS_2026_12_01)
        wb = load_xlsx_workbook_bytes(session.find_download(suffix=".xlsx")["bytes"])
        ws = wb.active
        headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
        hold_c = headers.index("Hold total time(Days)") + 1
        tid_c = headers.index("Ticket ID") + 1
        xlsx_hold = {ws.cell(r, tid_c).value: ws.cell(r, hold_c).value for r in range(2, ws.max_row + 1)}
        self.assertEqual(xlsx_hold, EXPECTED_HOLD_DELTAS_2026_12_01)

    def test_t3_03_filter_on_hold_then_change_today_updates_filtered_view(self) -> None:
        session = AppSession().start()
        raw_df = build_reference_9_rows_df("iso")
        session.upload_to(session.find_dashboard_uploader(), "Nov1-25.csv", to_csv_bytes(raw_df), "text/csv")
        session.set_date(session.find_today_date_input(), REFERENCE_TODAY)
        _apply_filter(session, "On Hold")
        kpi1 = session.extract_kpi_text()
        dfs1 = [
            d[["Ticket ID", "Hold total time(Days)"]].copy()
            if "Ticket ID" in d.columns
            else d[["Hold total time(Days)"]].copy()
            for d in session.get_dataframes()
            if "Hold total time(Days)" in d.columns
        ]
        charts1 = session.get_chart_signatures()

        session.set_date(session.find_today_date_input(), date(2026, 12, 5))
        session.assert_no_exceptions()
        kpi2 = session.extract_kpi_text()
        dfs2 = [
            d[["Ticket ID", "Hold total time(Days)"]].copy()
            if "Ticket ID" in d.columns
            else d[["Hold total time(Days)"]].copy()
            for d in session.get_dataframes()
            if "Hold total time(Days)" in d.columns
        ]
        charts2 = session.get_chart_signatures()

        kpi_changed = kpi1 != kpi2
        has_hold_kpi = any(
            "hold" in line and any(tok in line for tok in ("day", "avg", "max", "duration", "time"))
            for line in kpi1.lower().splitlines()
        )
        if has_hold_kpi:
            self.assertTrue(kpi_changed, "Hold duration KPI text should update when 'today' changes.")

        table_changed = len(dfs1) != len(dfs2) or any(
            not d1.equals(d2) for d1, d2 in zip(dfs1, dfs2)
        )
        charts_changed = charts1 != charts2
        self.assertTrue(
            kpi_changed or table_changed or charts_changed,
            "Changing 'today' under active On Hold filter should update KPIs, Hold table values, or charts.",
        )

        hold_tables = [d for d in session.get_dataframes() if "Ticket ID" in d.columns and len(d) == 2]
        self.assertTrue(hold_tables, "On Hold filter should remain active after changing today")
        self.assertEqual(set(hold_tables[0]["Ticket ID"]), {"EX-2026-46397", "EX-2026-40520"})
        if "Hold total time(Days)" in hold_tables[0].columns:
            by_tid = hold_tables[0].set_index("Ticket ID")
            self.assertEqual(int(by_tid.loc["EX-2026-46397", "Hold total time(Days)"]), 10)
            self.assertEqual(int(by_tid.loc["EX-2026-40520", "Hold total time(Days)"]), 20)

        out_df = parse_csv_bytes(session.find_download(suffix=".csv")["bytes"]).set_index("Ticket ID")
        self.assertEqual(int(out_df.loc["EX-2026-46397", "Hold total time(Days)"]), 10)
        self.assertEqual(int(out_df.loc["EX-2026-40520", "Hold total time(Days)"]), 20)

    def test_t3_04_filtered_dashboard_still_exports_valid_csv_and_xlsx(self) -> None:
        session = AppSession().start()
        raw_df = build_reference_9_rows_df("iso")
        session.upload_to(session.find_dashboard_uploader(), "Nov1-25.csv", to_csv_bytes(raw_df), "text/csv")
        session.set_date(session.find_today_date_input(), ALT_TODAY)
        _apply_filter(session, "Critical")
        csv_df = parse_csv_bytes(session.find_download(suffix=".csv")["bytes"])
        wb = load_xlsx_workbook_bytes(session.find_download(suffix=".xlsx")["bytes"])
        self.assertGreaterEqual(len(csv_df), 2)
        self.assertEqual(wb.active.max_row - 1, len(csv_df))

    def test_t3_05_view_switching_preserves_dashboard_and_compare_independently(self) -> None:
        session = AppSession().start()
        raw_df = build_reference_9_rows_df("iso")
        session.upload_to(session.find_dashboard_uploader(), "Nov1-25.csv", to_csv_bytes(raw_df), "text/csv")
        session.set_date(session.find_today_date_input(), REFERENCE_TODAY)

        session.switch_view(VIEW_OMS_FIR_COMPARE)
        session.click_button("Sample Files")
        session.click_button("Compare Both Files")
        self.assertIn("All Chargers List (54)", session.get_tab_labels())

        session.switch_view(VIEW_OMS_DASHBOARD)
        out_df = parse_csv_bytes(session.find_download(suffix=".csv")["bytes"])
        assert_reference_converted_df(out_df)

    def test_t3_06_xls_upload_filter_search_and_export_matches_xlsx(self) -> None:
        raw_df = build_reference_9_rows_df("iso")
        s_xls = AppSession().start()
        s_xls.upload_to(s_xls.find_dashboard_uploader(), "oms.xls", to_xls_xml_bytes(raw_df), "application/vnd.ms-excel")
        s_xls.set_date(s_xls.find_today_date_input(), REFERENCE_TODAY)
        _apply_filter(s_xls, "Critical")
        s_xls.set_text(s_xls.find_search_input(), "32559")
        s_xls.assert_no_exceptions()
        matched = [d for d in s_xls.get_dataframes() if "Ticket ID" in d.columns and len(d) == 1]
        self.assertTrue(matched)
        self.assertEqual(matched[0].iloc[0]["Ticket ID"], "EX-2026-32559")

        # Clear filters & search before asserting full-dataset CSV parity against unfiltered .xlsx export
        s_xls.clear_dashboard_filters()
        restored = [d for d in s_xls.get_dataframes() if "Ticket ID" in d.columns and len(d) == 9]
        self.assertTrue(restored, "Clearing filters & search should restore all 9 rows")
        xls_csv = s_xls.find_download(suffix=".csv")["bytes"]

        s_xlsx = AppSession().start()
        s_xlsx.upload_to(
            s_xlsx.find_dashboard_uploader(),
            "oms.xlsx",
            to_xlsx_bytes(raw_df),
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        s_xlsx.set_date(s_xlsx.find_today_date_input(), REFERENCE_TODAY)
        xlsx_csv = s_xlsx.find_download(suffix=".csv")["bytes"]
        pd.testing.assert_frame_equal(parse_csv_bytes(xls_csv), parse_csv_bytes(xlsx_csv))

    def test_t3_07_cross_midnight_negative_zero_round_trip_via_ui(self) -> None:
        s1 = AppSession().start()
        raw_df = build_reference_9_rows_df("iso")
        s1.upload_to(s1.find_dashboard_uploader(), "Nov1-25.csv", to_csv_bytes(raw_df), "text/csv")
        s1.set_date(s1.find_today_date_input(), REFERENCE_TODAY)
        exported_csv_bytes = s1.find_download(suffix=".csv")["bytes"]

        s2 = AppSession().start()
        s2.upload_to(s2.find_dashboard_uploader(), "Nov1-25_date_only.csv", exported_csv_bytes, "text/csv")
        s2.set_date(s2.find_today_date_input(), REFERENCE_TODAY)
        s2.assert_no_exceptions()
        reexported = parse_csv_bytes(s2.find_download(suffix=".csv")["bytes"])
        assert_reference_converted_df(reexported)

    def test_t3_08_subset_24_col_oms_dump_filter_search_and_export(self) -> None:
        self.assertTrue(DISK_OMS_PATH.exists())
        session = AppSession().start()
        session.upload_to(
            session.find_dashboard_uploader(),
            "OMS DUMP.xlsx",
            DISK_OMS_PATH.read_bytes(),
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        session.assert_no_exceptions()
        csv_df = parse_csv_bytes(session.find_download(suffix=".csv")["bytes"])
        self.assertEqual(len(csv_df), 2605)
        self.assertEqual(len(csv_df.columns), 24)
        _apply_filter(session, "North")
        session.set_text(session.find_search_input(), "EX-2026-4574")
        session.assert_no_exceptions()
        dfs = [d for d in session.get_dataframes() if "Ticket ID" in d.columns and len(d) == 1]
        self.assertTrue(dfs, "Expected 1 matching ticket EX-2026-4574 in North zone")

    def test_t3_09_reupload_file_b_inherits_active_today_and_replaces_file_a(self) -> None:
        session = AppSession().start()
        df_a = build_reference_9_rows_df("iso")
        session.upload_to(session.find_dashboard_uploader(), "file_a.csv", to_csv_bytes(df_a), "text/csv")
        session.set_date(session.find_today_date_input(), ALT_TODAY)

        df_b = df_a.iloc[[2, 4, 5]].copy()
        session.upload_to(session.find_dashboard_uploader(), "file_b.csv", to_csv_bytes(df_b), "text/csv")
        session.assert_no_exceptions()
        out_df = parse_csv_bytes(session.find_download(suffix=".csv")["bytes"])
        self.assertEqual(len(out_df), 3)
        by_tid = out_df.set_index("Ticket ID")
        self.assertEqual(int(by_tid.loc["EX-2026-46397", "Hold total time(Days)"]), 6)
        self.assertEqual(int(by_tid.loc["EX-2026-40520", "Hold total time(Days)"]), 16)
        self.assertEqual(int(by_tid.loc["EX-2026-32559", "Hold total time(Days)"]), 25)

    def test_t3_10_compare_filtered_west_missing_tickets_excel_and_csv_export(self) -> None:
        session = AppSession().start()
        session.switch_view(VIEW_OMS_FIR_COMPARE)
        session.click_button("Sample Files")
        session.click_button("Compare Both Files")
        session.select_pill_containing("Missing Tickets")
        session.set_multiselect(session.find_multiselect("Region / City"), ["West"])
        session.assert_no_exceptions()
        xlsx_dl = session.find_download(suffix=".xlsx")
        wb = load_xlsx_workbook_bytes(xlsx_dl["bytes"])
        self.assertIn("Ticket Generation Gaps", wb.sheetnames)
        csv_dl = session.find_download(filename_contains="missing_tickets")
        missing_west = parse_csv_bytes(csv_dl["bytes"])
        self.assertGreaterEqual(len(missing_west), 1)
        if "Zone" in missing_west.columns:
            self.assertEqual(set(missing_west["Zone"]), {"West"})

    def test_t3_11_double_space_resolved_column_flows_into_ui_charts_and_table(self) -> None:
        session = AppSession().start()
        raw_df = build_reference_9_rows_df("iso")
        session.upload_to(session.find_dashboard_uploader(), "Nov1-25.csv", to_csv_bytes(raw_df), "text/csv")
        session.set_date(session.find_today_date_input(), REFERENCE_TODAY)
        session.assert_no_exceptions()
        dfs = session.get_dataframes()
        self.assertTrue(any(DATE_ONLY_RESOLVED_COL in d.columns for d in dfs))
        self.assertGreaterEqual(len(session.get_charts()), 1)

    def test_t3_12_error_recovery_dashboard_corrupt_then_compare_then_valid_csv(self) -> None:
        session = AppSession().start()
        session.upload_to(
            session.find_dashboard_uploader(),
            "bad.xlsx",
            corrupt_binary_bytes(),
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        session.assert_no_exceptions()
        self.assertGreaterEqual(len(session.get_error_messages()), 1)

        session.switch_view(VIEW_OMS_FIR_COMPARE)
        session.click_button("Sample Files")
        session.click_button("Compare Both Files")
        self.assertIn("All Chargers List (54)", session.get_tab_labels())

        session.switch_view(VIEW_OMS_DASHBOARD)
        session.upload_to(
            session.find_dashboard_uploader(),
            "good.csv",
            to_csv_bytes(build_reference_9_rows_df("iso")),
            "text/csv",
        )
        session.set_date(session.find_today_date_input(), REFERENCE_TODAY)
        session.assert_no_exceptions()
        self.assertEqual(len(session.get_error_messages()), 0)
        assert_reference_converted_df(parse_csv_bytes(session.find_download(suffix=".csv")["bytes"]))

    def test_t3_13_contradictory_empty_filter_reset_and_export(self) -> None:
        session = AppSession().start()
        session.upload_to(
            session.find_dashboard_uploader(),
            "Nov1-25.csv",
            to_csv_bytes(build_reference_9_rows_df("iso")),
            "text/csv",
        )
        session.set_date(session.find_today_date_input(), REFERENCE_TODAY)
        _apply_filter(session, "Critical")
        session.set_text(session.find_search_input(), "EX-2026-34312")  # Low/Closed, not Critical
        session.assert_no_exceptions()
        self.assertRegex(session.extract_kpi_text(), r"\b0\b")
        session.set_text(session.find_search_input(), "")
        session.assert_no_exceptions()
        out_df = parse_csv_bytes(session.find_download(suffix=".csv")["bytes"])
        assert_reference_converted_df(out_df)

    def test_t3_14_row1_formula_hints_export_as_none_not_literal_strings(self) -> None:
        session = AppSession().start()
        session.upload_to(
            session.find_dashboard_uploader(),
            "Nov1-25.xlsx",
            to_xlsx_bytes(build_reference_9_rows_df("iso")),
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        session.set_date(session.find_today_date_input(), REFERENCE_TODAY)
        wb = load_xlsx_workbook_bytes(session.find_download(suffix=".xlsx")["bytes"])
        ws = wb.active
        headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
        for dcol in DELTA_COLS_8:
            val = ws.cell(2, headers.index(dcol) + 1).value
            self.assertIsNone(val, f"Row 1 {dcol} should be None, got {val!r}")

    def test_t3_15_shared_oms_dump_in_both_views_zero_cross_contamination(self) -> None:
        self.assertTrue(DISK_OMS_PATH.exists())
        session = AppSession().start()
        session.upload_to(
            session.find_dashboard_uploader(),
            "OMS DUMP.xlsx",
            DISK_OMS_PATH.read_bytes(),
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        session.assert_no_exceptions()
        session.switch_view(VIEW_OMS_FIR_COMPARE)
        session.click_button("Sample Files")
        session.click_button("Compare Both Files")
        self.assertIn("All Chargers List (54)", session.get_tab_labels())

        session.switch_view(VIEW_OMS_DASHBOARD)
        out_df = parse_csv_bytes(session.find_download(suffix=".csv")["bytes"])
        self.assertEqual(len(out_df), 2605)
        self.assertEqual(len(out_df.columns), 24)


if __name__ == "__main__":
    unittest.main()
