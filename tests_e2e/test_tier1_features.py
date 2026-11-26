"""Tier 1 E2E Tests: Core Feature Happy Paths (55 tests = 5 per feature F1-F11)."""

from __future__ import annotations

import csv
from datetime import date
import io
import re
import unittest

import pandas as pd

from tests_e2e.e2e_harness import (
    AppSession,
    VIEW_OMS_DASHBOARD,
    VIEW_OMS_FIR_COMPARE,
    assert_reference_converted_df,
    get_core_loader,
    get_oms_converter,
    load_xlsx_workbook_bytes,
    parse_csv_bytes,
)
from tests_e2e.fixtures import (
    ALT_TODAY,
    DATE_ONLY_RESOLVED_COL,
    DATETIME_COLS_15,
    DELTA_COLS_8,
    DISK_OMS_PATH,
    EXPECTED_DELTAS_2026_11_25,
    EXPECTED_HOLD_DELTAS_2026_12_01,
    EXPECTED_STRIPPED_DATES,
    RAW_RESOLVED_COL,
    REFERENCE_42_COLUMNS_DATE_ONLY,
    REFERENCE_TODAY,
    build_reference_9_rows_df,
    to_csv_bytes,
    to_xls_xml_bytes,
    to_xlsx_bytes,
)


def _upload_ref_csv_to_dashboard(session: AppSession, *, set_today: date | None = REFERENCE_TODAY) -> AppSession:
    df = build_reference_9_rows_df("iso")
    up = session.find_dashboard_uploader()
    session.upload_to(up, "Nov1-25-OMS.csv", to_csv_bytes(df), "text/csv")
    session.assert_no_exceptions()
    if set_today is not None:
        di = session.find_today_date_input()
        if di.value != set_today:
            session.set_date(di, set_today)
            session.assert_no_exceptions()
    return session


def _filter_dashboard_by_option(session: AppSession, option_value: str) -> None:
    """Apply a filter option via pills, multiselect, or selectbox matching option_value."""
    session.select_dashboard_filter(option_value)


class TestTier1F1Navigation(unittest.TestCase):
    def test_t1_f1_01_default_view_is_oms_dashboard(self) -> None:
        session = AppSession().start()
        session.assert_no_exceptions()
        nav = session.get_nav_radio()
        opts = [str(o) for o in nav.options]
        self.assertEqual(opts[0], VIEW_OMS_DASHBOARD)
        self.assertEqual(opts[1], VIEW_OMS_FIR_COMPARE)
        self.assertEqual(nav.value, VIEW_OMS_DASHBOARD)

    def test_t1_f1_02_switch_to_oms_fir_compare(self) -> None:
        session = AppSession().start()
        session.switch_view(VIEW_OMS_FIR_COMPARE)
        text = session.extract_all_text()
        self.assertIn("EV Charger Complaint & Status Checker", text)
        btn = session.find_button("Sample Files")
        self.assertIsNotNone(btn)

    def test_t1_f1_03_switch_back_to_oms_dashboard(self) -> None:
        session = AppSession().start()
        session.switch_view(VIEW_OMS_FIR_COMPARE)
        session.switch_view(VIEW_OMS_DASHBOARD)
        text = session.extract_all_text()
        self.assertNotIn("EV Charger Complaint & Status Checker", text)
        _ = session.find_dashboard_uploader()
        _ = session.find_today_date_input()

    def test_t1_f1_04_nav_remains_accessible_after_dashboard_upload(self) -> None:
        session = AppSession().start()
        _upload_ref_csv_to_dashboard(session)
        nav = session.get_nav_radio()
        self.assertEqual(list(nav.options)[:2], [VIEW_OMS_DASHBOARD, VIEW_OMS_FIR_COMPARE])
        session.switch_view(VIEW_OMS_FIR_COMPARE)
        self.assertIn("EV Charger Complaint & Status Checker", session.extract_all_text())

    def test_t1_f1_05_nav_remains_accessible_after_compare_run(self) -> None:
        session = AppSession().start()
        session.switch_view(VIEW_OMS_FIR_COMPARE)
        session.click_button("Sample Files")
        session.click_button("Compare Both Files")
        session.assert_no_exceptions()
        nav = session.get_nav_radio()
        self.assertEqual(list(nav.options)[:2], [VIEW_OMS_DASHBOARD, VIEW_OMS_FIR_COMPARE])
        session.switch_view(VIEW_OMS_DASHBOARD)
        self.assertEqual(session.get_nav_radio().value, VIEW_OMS_DASHBOARD)


class TestTier1F2Isolation(unittest.TestCase):
    def test_t1_f2_01_distinct_uploaders_per_view(self) -> None:
        session = AppSession().start()
        dash_up = session.find_dashboard_uploader()
        dash_key = getattr(dash_up, "key", None)
        session.switch_view(VIEW_OMS_FIR_COMPARE)
        fir_up, oms_up = session.find_compare_uploaders()
        self.assertNotEqual(dash_key, getattr(fir_up, "key", None))
        self.assertNotEqual(dash_key, getattr(oms_up, "key", None))

    def test_t1_f2_02_dashboard_upload_does_not_populate_compare(self) -> None:
        session = AppSession().start()
        _upload_ref_csv_to_dashboard(session)
        session.switch_view(VIEW_OMS_FIR_COMPARE)
        text = session.extract_all_text()
        self.assertIn("Not uploaded yet", text)
        compare_btn = session.find_button("Compare Both Files")
        self.assertTrue(compare_btn.disabled)

    def test_t1_f2_03_compare_sample_does_not_leak_to_dashboard(self) -> None:
        session = AppSession().start()
        session.switch_view(VIEW_OMS_FIR_COMPARE)
        session.click_button("Sample Files")
        session.click_button("Compare Both Files")
        session.assert_no_exceptions()
        session.switch_view(VIEW_OMS_DASHBOARD)
        text = session.extract_all_text()
        self.assertNotIn("Charger & Ticket Check Results", text)
        self.assertEqual(len(session.get_downloads()), 0)

    def test_t1_f2_04_compare_clear_all_preserves_dashboard_state(self) -> None:
        session = AppSession().start()
        _upload_ref_csv_to_dashboard(session, set_today=ALT_TODAY)
        session.switch_view(VIEW_OMS_FIR_COMPARE)
        session.click_button("Sample Files")
        session.click_button("Clear All")
        session.assert_no_exceptions()
        session.switch_view(VIEW_OMS_DASHBOARD)
        self.assertEqual(session.find_today_date_input().value, ALT_TODAY)
        csv_dl = session.find_download(suffix=".csv")
        out_df = parse_csv_bytes(csv_dl["bytes"])
        self.assertEqual(len(out_df), 9)

    def test_t1_f2_05_distinct_datasets_active_simultaneously(self) -> None:
        session = AppSession().start()
        _upload_ref_csv_to_dashboard(session)
        session.switch_view(VIEW_OMS_FIR_COMPARE)
        session.click_button("Sample Files")
        session.click_button("Compare Both Files")
        self.assertIn("478", session.extract_kpi_text() + "\n" + " ".join(session.get_tab_labels()))
        session.switch_view(VIEW_OMS_DASHBOARD)
        csv_dl = session.find_download(suffix=".csv")
        out_df = parse_csv_bytes(csv_dl["bytes"])
        self.assertEqual(len(out_df), 9)
        self.assertIn("EX-2026-32559", set(out_df["Ticket ID"]))


class TestTier1F3CompareRegression(unittest.TestCase):
    def test_t1_f3_01_sample_files_status_and_enablement(self) -> None:
        session = AppSession().start()
        session.switch_view(VIEW_OMS_FIR_COMPARE)
        session.click_button("Sample Files")
        session.assert_no_exceptions()
        all_txt = session.extract_all_text()
        self.assertIn("54 chargers found", all_txt)
        self.assertIn("2605 tickets found", all_txt)
        self.assertFalse(session.find_button("Compare Both Files").disabled)

    def test_t1_f3_02_compare_both_files_baseline_kpis_and_tabs(self) -> None:
        session = AppSession().start()
        session.switch_view(VIEW_OMS_FIR_COMPARE)
        session.click_button("Sample Files")
        session.click_button("Compare Both Files")
        session.assert_no_exceptions()
        self.assertIn("Charger & Ticket Check Results", session.extract_all_text())
        kpi = session.extract_kpi_text()
        for num in ("54", "38", "5", "13", "20", "47.4%"):
            self.assertIn(num, kpi)
        tabs = session.get_tab_labels()
        self.assertEqual(
            tabs,
            [
                "Missing Tickets (5)",
                "Wrongly Closed Tickets (13)",
                "Being Worked On (20)",
                "Charts & Trends",
                "Extra Tickets in OMS (478)",
                "All Chargers List (54)",
            ],
        )

    def test_t1_f3_03_quick_preset_pill_and_search(self) -> None:
        session = AppSession().start()
        session.switch_view(VIEW_OMS_FIR_COMPARE)
        session.click_button("Sample Files")
        session.click_button("Compare Both Files")
        session.select_pill_containing("Wrongly Closed")
        session.assert_no_exceptions()
        pill_tabs = session.get_tab_labels()
        self.assertIn("Wrongly Closed Tickets (13)", pill_tabs)
        self.assertIn("Missing Tickets (0)", pill_tabs)
        ti = session.find_search_input()
        session.set_text(ti, "JH0007")
        session.assert_no_exceptions()
        dfs = session.get_dataframes()
        self.assertTrue(any(len(df) == 1 and "JH0007" in df.to_string() for df in dfs))

    def test_t1_f3_04_cascading_zone_filter_and_reset(self) -> None:
        session = AppSession().start()
        session.switch_view(VIEW_OMS_FIR_COMPARE)
        session.click_button("Sample Files")
        session.click_button("Compare Both Files")
        zone_ms = session.find_multiselect("Region / City")
        session.set_multiselect(zone_ms, ["West"])
        session.assert_no_exceptions()
        self.assertIn("All Chargers List (", " ".join(session.get_tab_labels()))
        self.assertNotIn("All Chargers List (54)", session.get_tab_labels())
        session.click_button("Reset Filters")
        session.assert_no_exceptions()
        self.assertIn("All Chargers List (54)", session.get_tab_labels())

    def test_t1_f3_05_excel_and_csv_downloads_and_clear_all(self) -> None:
        session = AppSession().start()
        session.switch_view(VIEW_OMS_FIR_COMPARE)
        session.click_button("Sample Files")
        session.click_button("Compare Both Files")
        xlsx_dl = session.find_download(suffix=".xlsx")
        wb = load_xlsx_workbook_bytes(xlsx_dl["bytes"])
        self.assertIn("Executive Summary", wb.sheetnames)
        self.assertEqual(len(wb.sheetnames), 6)
        csv_dl = session.find_download(filename_contains="missing_tickets")
        missing_df = parse_csv_bytes(csv_dl["bytes"])
        self.assertEqual(len(missing_df), 5)
        session.click_button("Clear All")
        session.assert_no_exceptions()
        self.assertIn("EV Charger Complaint & Status Checker", session.extract_all_text())


class TestTier1F4Loader(unittest.TestCase):
    def test_t1_f4_01_upload_csv_in_oms_dashboard(self) -> None:
        session = AppSession().start()
        _upload_ref_csv_to_dashboard(session)
        csv_dl = session.find_download(suffix=".csv")
        out_df = parse_csv_bytes(csv_dl["bytes"])
        assert_reference_converted_df(out_df)

    def test_t1_f4_02_upload_xlsx_in_oms_dashboard(self) -> None:
        session = AppSession().start()
        raw_df = build_reference_9_rows_df("iso")
        up = session.find_dashboard_uploader()
        session.upload_to(
            up,
            "Nov1-25-OMS.xlsx",
            to_xlsx_bytes(raw_df),
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        session.set_date(session.find_today_date_input(), REFERENCE_TODAY)
        session.assert_no_exceptions()
        out_df = parse_csv_bytes(session.find_download(suffix=".csv")["bytes"])
        assert_reference_converted_df(out_df)

    def test_t1_f4_03_upload_xls_xml_in_oms_dashboard(self) -> None:
        raw_df = build_reference_9_rows_df("iso")
        xls_bytes = to_xls_xml_bytes(raw_df)
        conv = get_oms_converter()
        loaded = conv.load_oms_raw_file(xls_bytes, "Nov1-25-OMS.xls")
        self.assertEqual(len(loaded), 9)
        session = AppSession().start()
        up = session.find_dashboard_uploader()
        session.upload_to(up, "Nov1-25-OMS.xls", xls_bytes, "application/vnd.ms-excel")
        session.set_date(session.find_today_date_input(), REFERENCE_TODAY)
        session.assert_no_exceptions()
        out_df = parse_csv_bytes(session.find_download(suffix=".csv")["bytes"])
        assert_reference_converted_df(out_df)

    def test_t1_f4_04_load_utf8_bom_csv(self) -> None:
        raw_df = build_reference_9_rows_df("iso")
        bom_bytes = to_csv_bytes(raw_df, bom=True)
        conv = get_oms_converter()
        loaded = conv.load_oms_raw_file(bom_bytes, "Nov1-25-OMS-bom.csv")
        self.assertEqual(list(loaded.columns)[0], "Sno")
        self.assertNotIn("\ufeffSno", loaded.columns)
        self.assertEqual(len(loaded), 9)

    def test_t1_f4_05_load_xlsx_with_excel_serial_dates(self) -> None:
        raw_df = build_reference_9_rows_df("iso")
        xlsx_bytes = to_xlsx_bytes(
            raw_df,
            excel_serial_cols=("Created At", "Assigned At", "Hold Status Time", "Resolved At Time"),
        )
        conv = get_oms_converter()
        loaded = conv.load_oms_raw_file(xlsx_bytes, "serial_dates.xlsx")
        out_df = conv.convert_oms_to_date_only(loaded, today_date=REFERENCE_TODAY)
        assert_reference_converted_df(out_df)


class TestTier1F5DateStrip(unittest.TestCase):
    def test_t1_f5_01_strip_iso_timestamps_across_all_15_cols(self) -> None:
        conv = get_oms_converter()
        out_df = conv.convert_oms_to_date_only(build_reference_9_rows_df("iso"), today_date=REFERENCE_TODAY)
        by_tid = out_df.set_index("Ticket ID")
        for tid, expected_map in EXPECTED_STRIPPED_DATES.items():
            for col in DATETIME_COLS_15:
                val = by_tid.loc[tid, col]
                actual = "" if pd.isna(val) else str(val).strip()
                self.assertEqual(actual, expected_map[col], f"{tid}.{col}")
                if actual:
                    self.assertRegex(actual, r"^\d{4}-\d{2}-\d{2}$")

    def test_t1_f5_02_strip_iso_z_and_offset_timestamps(self) -> None:
        conv = get_oms_converter()
        self.assertEqual(conv.strip_to_date_str("2026-11-23T14:30:00Z"), "2026-11-23")
        self.assertEqual(conv.strip_to_date_str("2026-11-23T23:45:00+05:30"), "2026-11-23")

    def test_t1_f5_03_strip_day_first_ddmmyyyy_strings(self) -> None:
        conv = get_oms_converter()
        self.assertEqual(conv.strip_to_date_str("25-11-2026 08:00"), "2026-11-25")
        self.assertEqual(conv.strip_to_date_str("19/11/2026"), "2026-11-19")
        self.assertEqual(conv.strip_to_date_str("03-11-2026 17:45:00"), "2026-11-03")

    def test_t1_f5_04_strip_named_month_12h_timestamps(self) -> None:
        conv = get_oms_converter()
        self.assertEqual(conv.strip_to_date_str("19 Nov 2026, 05:15:10 PM"), "2026-11-19")
        self.assertEqual(conv.strip_to_date_str("04 Nov 2026, 11:00:00 AM"), "2026-11-04")

    def test_t1_f5_05_strip_excel_serial_numbers(self) -> None:
        conv = get_oms_converter()
        self.assertEqual(conv.strip_to_date_str(46345), "2026-11-19")
        self.assertEqual(conv.strip_to_date_str(46351.75), "2026-11-25")


class TestTier1F6Rename(unittest.TestCase):
    def test_t1_f6_01_rename_preserves_double_space(self) -> None:
        conv = get_oms_converter()
        out_df = conv.convert_oms_to_date_only(build_reference_9_rows_df("iso"), today_date=REFERENCE_TODAY)
        self.assertIn(DATE_ONLY_RESOLVED_COL, out_df.columns)
        self.assertIn("  ", DATE_ONLY_RESOLVED_COL)

    def test_t1_f6_02_rename_stays_at_exact_column_position_28(self) -> None:
        conv = get_oms_converter()
        out_df = conv.convert_oms_to_date_only(build_reference_9_rows_df("iso"), today_date=REFERENCE_TODAY)
        self.assertEqual(list(out_df.columns)[27], DATE_ONLY_RESOLVED_COL)

    def test_t1_f6_03_raw_days_hours_header_removed(self) -> None:
        conv = get_oms_converter()
        out_df = conv.convert_oms_to_date_only(build_reference_9_rows_df("iso"), today_date=REFERENCE_TODAY)
        self.assertNotIn(RAW_RESOLVED_COL, out_df.columns)

    def test_t1_f6_04_idempotent_when_already_renamed(self) -> None:
        conv = get_oms_converter()
        first = conv.convert_oms_to_date_only(build_reference_9_rows_df("iso"), today_date=REFERENCE_TODAY)
        second = conv.convert_oms_to_date_only(first, today_date=REFERENCE_TODAY)
        self.assertEqual(list(second.columns), list(REFERENCE_42_COLUMNS_DATE_ONLY))
        self.assertEqual(list(second.columns).count(DATE_ONLY_RESOLVED_COL), 1)

    def test_t1_f6_05_renamed_col_contains_computed_integer_deltas(self) -> None:
        conv = get_oms_converter()
        out_df = conv.convert_oms_to_date_only(build_reference_9_rows_df("iso"), today_date=REFERENCE_TODAY)
        by_tid = out_df.set_index("Ticket ID")
        self.assertEqual(int(by_tid.loc["IT-2026-49625", DATE_ONLY_RESOLVED_COL]), 1)
        self.assertEqual(int(by_tid.loc["EX-2026-32559", DATE_ONLY_RESOLVED_COL]), 18)
        self.assertEqual(int(by_tid.loc["EX-2026-46325", DATE_ONLY_RESOLVED_COL]), 3)
        self.assertTrue(pd.isna(by_tid.loc["EX-2026-49803", DATE_ONLY_RESOLVED_COL]))


class TestTier1F7Schema(unittest.TestCase):
    def test_t1_f7_01_exact_42_column_order_preserved(self) -> None:
        conv = get_oms_converter()
        out_df = conv.convert_oms_to_date_only(build_reference_9_rows_df("iso"), today_date=REFERENCE_TODAY)
        self.assertEqual(len(out_df.columns), 42)
        self.assertEqual(list(out_df.columns), list(REFERENCE_42_COLUMNS_DATE_ONLY))

    def test_t1_f7_02_confirmed_positions_match_colomn_indices(self) -> None:
        conv = get_oms_converter()
        out_df = conv.convert_oms_to_date_only(build_reference_9_rows_df("iso"), today_date=REFERENCE_TODAY)
        cols = list(out_df.columns)
        self.assertEqual(cols[18], "Assigned At")
        self.assertEqual(cols[20], "Acknowledged On")
        self.assertEqual(cols[24], "Hold Status Time")
        self.assertEqual(cols[26], "Resolved At Time")
        self.assertEqual(cols[27], DATE_ONLY_RESOLVED_COL)
        self.assertEqual(cols[28], "Restored At Time")
        self.assertEqual(cols[30], "Closed At")
        self.assertEqual(cols[32], "Resolution SLA Deadline")
        self.assertEqual(cols[35], "Approval Requested At Time")
        self.assertEqual(cols[36], "Approval Accepted At Time")
        self.assertEqual(cols[38], "CET Approval Requested At Time")
        self.assertEqual(cols[39], "CET Approval Accepted At Time")

    def test_t1_f7_03_metadata_columns_pass_through_unmodified(self) -> None:
        raw_df = build_reference_9_rows_df("iso")
        conv = get_oms_converter()
        out_df = conv.convert_oms_to_date_only(raw_df, today_date=REFERENCE_TODAY)
        for col in ("Sno", "Ticket ID", "Station", "Zone", "Status", "Priority", "Charger", "Remarks"):
            self.assertEqual(list(out_df[col]), list(raw_df[col]))

    def test_t1_f7_04_input_dataframe_not_mutated_in_place(self) -> None:
        raw_df = build_reference_9_rows_df("iso")
        snapshot = raw_df.copy(deep=True)
        conv = get_oms_converter()
        _ = conv.convert_oms_to_date_only(raw_df, today_date=REFERENCE_TODAY)
        self.assertEqual(list(raw_df.columns), list(snapshot.columns))
        pd.testing.assert_frame_equal(raw_df, snapshot)

    def test_t1_f7_05_subset_24_col_oms_dump_handled_gracefully(self) -> None:
        self.assertTrue(DISK_OMS_PATH.exists(), f"Missing disk sample: {DISK_OMS_PATH}")
        conv = get_oms_converter()
        raw_df = conv.load_oms_raw_file(DISK_OMS_PATH, "OMS DUMP.xlsx")
        out_df = conv.convert_oms_to_date_only(raw_df, today_date=REFERENCE_TODAY)
        self.assertEqual(len(out_df), 2605)
        self.assertEqual(list(out_df.columns), list(raw_df.columns))
        row0 = out_df.loc[out_df["Ticket ID"] == "EX-2026-4574"].iloc[0]
        self.assertEqual(str(row0["Created At"]), "2026-09-30")
        self.assertEqual(str(row0["Updated At"]), "2026-11-19")
        self.assertEqual(str(row0["Closed At"]), "2026-11-19")


class TestTier1F8Deltas(unittest.TestCase):
    def test_t1_f8_01_acknowledgement_time_taken_days(self) -> None:
        conv = get_oms_converter()
        out_df = conv.convert_oms_to_date_only(build_reference_9_rows_df("iso"), today_date=REFERENCE_TODAY)
        by_tid = out_df.set_index("Ticket ID")
        for tid, exp in EXPECTED_DELTAS_2026_11_25.items():
            val = by_tid.loc[tid, "Acknowledgement time taken(Days)"]
            if exp["Acknowledgement time taken(Days)"] is None:
                self.assertTrue(pd.isna(val), tid)
            else:
                self.assertEqual(int(val), exp["Acknowledgement time taken(Days)"], tid)

    def test_t1_f8_02_hold_total_time_days_for_reference_today(self) -> None:
        conv = get_oms_converter()
        out_df = conv.convert_oms_to_date_only(build_reference_9_rows_df("iso"), today_date=REFERENCE_TODAY)
        by_tid = out_df.set_index("Ticket ID")
        for tid, exp in EXPECTED_DELTAS_2026_11_25.items():
            val = by_tid.loc[tid, "Hold total time(Days)"]
            if exp["Hold total time(Days)"] is None:
                self.assertTrue(pd.isna(val), tid)
            else:
                self.assertEqual(int(val), exp["Hold total time(Days)"], tid)

    def test_t1_f8_03_resolved_restore_and_total_time_days(self) -> None:
        conv = get_oms_converter()
        out_df = conv.convert_oms_to_date_only(build_reference_9_rows_df("iso"), today_date=REFERENCE_TODAY)
        by_tid = out_df.set_index("Ticket ID")
        for tid, exp in EXPECTED_DELTAS_2026_11_25.items():
            for col in (DATE_ONLY_RESOLVED_COL, "Restore total time(Days)", "Total time taken (Days)"):
                val = by_tid.loc[tid, col]
                if exp[col] is None:
                    self.assertTrue(pd.isna(val), f"{tid}.{col}")
                else:
                    self.assertEqual(int(val), exp[col], f"{tid}.{col}")

    def test_t1_f8_04_sla_actual_approval_and_cet_approval_deltas(self) -> None:
        conv = get_oms_converter()
        out_df = conv.convert_oms_to_date_only(build_reference_9_rows_df("iso"), today_date=REFERENCE_TODAY)
        by_tid = out_df.set_index("Ticket ID")
        for tid, exp in EXPECTED_DELTAS_2026_11_25.items():
            for col in (
                "SLA Breached As per Actual (Days)",
                "Approval Time taken (Days)",
                "CET Approval Time taken (Days)",
            ):
                val = by_tid.loc[tid, col]
                if exp[col] is None:
                    self.assertTrue(pd.isna(val), f"{tid}.{col}")
                else:
                    self.assertEqual(int(val), exp[col], f"{tid}.{col}")

    def test_t1_f8_05_date_only_subtraction_before_delta(self) -> None:
        conv = get_oms_converter()
        df = build_reference_9_rows_df("iso").iloc[[0]].copy()
        df.loc[:, "Assigned At"] = "2026-11-10 22:15:00"
        df.loc[:, "Acknowledged On"] = "2026-11-11 04:00:00"
        out_df = conv.convert_oms_to_date_only(df, today_date=REFERENCE_TODAY)
        self.assertEqual(int(out_df.iloc[0]["Acknowledgement time taken(Days)"]), 1)


class TestTier1F9TodayControl(unittest.TestCase):
    def test_t1_f9_01_today_control_defaults_to_system_date(self) -> None:
        session = AppSession().start()
        session.assert_no_exceptions()
        di = session.find_today_date_input()
        self.assertEqual(di.value, date.today())

    def test_t1_f9_02_today_set_to_2026_11_25_matches_reference_hold(self) -> None:
        session = AppSession().start()
        _upload_ref_csv_to_dashboard(session, set_today=REFERENCE_TODAY)
        out_df = parse_csv_bytes(session.find_download(suffix=".csv")["bytes"])
        assert_reference_converted_df(out_df)

    def test_t1_f9_03_today_changed_forward_recalculates_hold_only(self) -> None:
        session = AppSession().start()
        _upload_ref_csv_to_dashboard(session, set_today=date(2026, 11, 30))
        out_df = parse_csv_bytes(session.find_download(suffix=".csv")["bytes"])
        expected_hold = {
            "EX-2026-49803": None,
            "EX-2026-42245": None,
            "EX-2026-46397": 5,
            "IT-2026-49625": None,
            "EX-2026-40520": 15,
            "EX-2026-32559": 24,
            "EX-2026-34312": None,
            "EX-2026-46325": None,
            "EX-2026-43053": None,
        }
        assert_reference_converted_df(out_df, expected_hold_map=expected_hold)

    def test_t1_f9_04_today_changed_backward_recalculates_hold(self) -> None:
        session = AppSession().start()
        _upload_ref_csv_to_dashboard(session, set_today=date(2026, 11, 20))
        out_df = parse_csv_bytes(session.find_download(suffix=".csv")["bytes"])
        expected_hold = {
            "EX-2026-49803": None,
            "EX-2026-42245": None,
            "EX-2026-46397": -5,
            "IT-2026-49625": None,
            "EX-2026-40520": 5,
            "EX-2026-32559": 14,
            "EX-2026-34312": None,
            "EX-2026-46325": None,
            "EX-2026-43053": None,
        }
        assert_reference_converted_df(out_df, expected_hold_map=expected_hold)

    def test_t1_f9_05_today_change_updates_ui_hold_kpi_text(self) -> None:
        session = AppSession().start()
        _upload_ref_csv_to_dashboard(session, set_today=REFERENCE_TODAY)
        kpi_nov25 = session.extract_kpi_text()
        dfs_nov25 = [
            d[["Ticket ID", "Hold total time(Days)"]].copy()
            if "Ticket ID" in d.columns
            else d[["Hold total time(Days)"]].copy()
            for d in session.get_dataframes()
            if "Hold total time(Days)" in d.columns
        ]
        charts_nov25 = session.get_chart_signatures()

        session.set_date(session.find_today_date_input(), date(2026, 12, 5))
        session.assert_no_exceptions()

        kpi_dec05 = session.extract_kpi_text()
        dfs_dec05 = [
            d[["Ticket ID", "Hold total time(Days)"]].copy()
            if "Ticket ID" in d.columns
            else d[["Hold total time(Days)"]].copy()
            for d in session.get_dataframes()
            if "Hold total time(Days)" in d.columns
        ]
        charts_dec05 = session.get_chart_signatures()

        kpi_changed = kpi_nov25 != kpi_dec05
        has_hold_kpi = any(
            "hold" in line and any(tok in line for tok in ("day", "avg", "max", "duration", "time"))
            for line in kpi_nov25.lower().splitlines()
        )
        if has_hold_kpi:
            self.assertTrue(kpi_changed, "Hold duration KPI text should update when 'today' changes.")

        table_changed = len(dfs_nov25) != len(dfs_dec05) or any(
            not d1.equals(d2) for d1, d2 in zip(dfs_nov25, dfs_dec05)
        )
        charts_changed = charts_nov25 != charts_dec05
        self.assertTrue(
            kpi_changed or table_changed or charts_changed,
            "Changing 'today' should update KPI text, rendered Hold table values, or charts.",
        )

        for d in dfs_dec05:
            if "Ticket ID" in d.columns and len(d) == 9:
                by_tid = d.set_index("Ticket ID")
                self.assertEqual(int(by_tid.loc["EX-2026-46397", "Hold total time(Days)"]), 10)
                self.assertEqual(int(by_tid.loc["EX-2026-40520", "Hold total time(Days)"]), 20)
                self.assertEqual(int(by_tid.loc["EX-2026-32559", "Hold total time(Days)"]), 29)

        out_df = parse_csv_bytes(session.find_download(suffix=".csv")["bytes"]).set_index("Ticket ID")
        self.assertEqual(int(out_df.loc["EX-2026-46397", "Hold total time(Days)"]), 10)
        self.assertEqual(int(out_df.loc["EX-2026-40520", "Hold total time(Days)"]), 20)
        self.assertEqual(int(out_df.loc["EX-2026-32559", "Hold total time(Days)"]), 29)


class TestTier1F10Export(unittest.TestCase):
    def test_t1_f10_01_dashboard_csv_download_valid_42_cols(self) -> None:
        session = AppSession().start()
        _upload_ref_csv_to_dashboard(session, set_today=REFERENCE_TODAY)
        csv_dl = session.find_download(suffix=".csv")
        out_df = parse_csv_bytes(csv_dl["bytes"])
        self.assertEqual(list(out_df.columns), list(REFERENCE_42_COLUMNS_DATE_ONLY))
        self.assertEqual(len(out_df), 9)

    def test_t1_f10_02_csv_raw_text_has_no_dot_zero_on_ints(self) -> None:
        session = AppSession().start()
        _upload_ref_csv_to_dashboard(session, set_today=REFERENCE_TODAY)
        csv_text = session.find_download(suffix=".csv")["bytes"].decode("utf-8")
        rows = {r["Ticket ID"]: r for r in csv.DictReader(io.StringIO(csv_text))}
        r_32559 = rows["EX-2026-32559"]
        self.assertEqual(r_32559["Acknowledgement time taken(Days)"], "0")
        self.assertEqual(r_32559["Hold total time(Days)"], "19")
        self.assertEqual(r_32559[DATE_ONLY_RESOLVED_COL], "18")
        self.assertEqual(r_32559["Approval Time taken (Days)"], "-18")
        self.assertEqual(r_32559["Total time taken (Days)"], "")
        r_49803 = rows["EX-2026-49803"]
        for dcol in DELTA_COLS_8:
            self.assertEqual(r_49803[dcol], "")

    def test_t1_f10_03_dashboard_xlsx_download_valid_workbook(self) -> None:
        session = AppSession().start()
        _upload_ref_csv_to_dashboard(session, set_today=REFERENCE_TODAY)
        xlsx_dl = session.find_download(suffix=".xlsx")
        wb = load_xlsx_workbook_bytes(xlsx_dl["bytes"])
        ws = wb.active
        headers = [ws.cell(row=1, column=c).value for c in range(1, ws.max_column + 1)]
        self.assertEqual(headers, list(REFERENCE_42_COLUMNS_DATE_ONLY))
        self.assertEqual(ws.max_row, 10)  # 1 header + 9 data rows

    def test_t1_f10_04_xlsx_cell_types_ints_and_none_blanks(self) -> None:
        session = AppSession().start()
        _upload_ref_csv_to_dashboard(session, set_today=REFERENCE_TODAY)
        wb = load_xlsx_workbook_bytes(session.find_download(suffix=".xlsx")["bytes"])
        ws = wb.active
        headers = [ws.cell(row=1, column=c).value for c in range(1, ws.max_column + 1)]
        col_idx = {name: i + 1 for i, name in enumerate(headers)}
        row_by_tid = {ws.cell(row=r, column=col_idx["Ticket ID"]).value: r for r in range(2, ws.max_row + 1)}

        r6 = row_by_tid["EX-2026-32559"]
        ack_val = ws.cell(row=r6, column=col_idx["Acknowledgement time taken(Days)"]).value
        app_val = ws.cell(row=r6, column=col_idx["Approval Time taken (Days)"]).value
        tot_val = ws.cell(row=r6, column=col_idx["Total time taken (Days)"]).value
        self.assertIs(type(ack_val), int)
        self.assertEqual(ack_val, 0)
        self.assertIs(type(app_val), int)
        self.assertEqual(app_val, -18)
        self.assertIsNone(tot_val)

        r1 = row_by_tid["EX-2026-49803"]
        for dcol in DELTA_COLS_8:
            self.assertIsNone(ws.cell(row=r1, column=col_idx[dcol]).value)

    def test_t1_f10_05_exports_reflect_updated_today_2026_12_01(self) -> None:
        session = AppSession().start()
        _upload_ref_csv_to_dashboard(session, set_today=ALT_TODAY)
        csv_df = parse_csv_bytes(session.find_download(suffix=".csv")["bytes"])
        assert_reference_converted_df(csv_df, expected_hold_map=EXPECTED_HOLD_DELTAS_2026_12_01)

        wb = load_xlsx_workbook_bytes(session.find_download(suffix=".xlsx")["bytes"])
        ws = wb.active
        headers = [ws.cell(row=1, column=c).value for c in range(1, ws.max_column + 1)]
        tid_c = headers.index("Ticket ID") + 1
        hold_c = headers.index("Hold total time(Days)") + 1
        actual_hold = {
            ws.cell(row=r, column=tid_c).value: ws.cell(row=r, column=hold_c).value
            for r in range(2, ws.max_row + 1)
        }
        self.assertEqual(actual_hold, EXPECTED_HOLD_DELTAS_2026_12_01)


class TestTier1F11DashboardUI(unittest.TestCase):
    def test_t1_f11_01_kpi_cards_render_for_9_rows(self) -> None:
        session = AppSession().start()
        _upload_ref_csv_to_dashboard(session)
        kpi_text = session.extract_kpi_text()
        self.assertTrue(len(kpi_text.strip()) > 0, "Expected KPI cards rendered on OMS Dashboard")
        self.assertRegex(kpi_text, r"\b9\b")

    def test_t1_f11_02_status_or_preset_filter_narrows_table_and_kpis(self) -> None:
        session = AppSession().start()
        _upload_ref_csv_to_dashboard(session)
        _filter_dashboard_by_option(session, "On Hold")
        dfs = session.get_dataframes()
        matched = [df for df in dfs if "Ticket ID" in df.columns and len(df) == 2]
        self.assertTrue(matched, f"Expected filtered 2-row table for On Hold; got {[len(d) for d in dfs]}")
        self.assertEqual(set(matched[0]["Ticket ID"]), {"EX-2026-46397", "EX-2026-40520"})

    def test_t1_f11_03_priority_filter_narrows_to_critical(self) -> None:
        session = AppSession().start()
        _upload_ref_csv_to_dashboard(session)
        _filter_dashboard_by_option(session, "Critical")
        dfs = session.get_dataframes()
        matched = [df for df in dfs if "Ticket ID" in df.columns and len(df) == 2]
        self.assertTrue(matched, f"Expected filtered 2-row table for Critical; got {[len(d) for d in dfs]}")
        self.assertEqual(set(matched[0]["Ticket ID"]), {"EX-2026-46397", "EX-2026-32559"})

    def test_t1_f11_04_live_search_filters_and_clears(self) -> None:
        session = AppSession().start()
        _upload_ref_csv_to_dashboard(session)
        ti = session.find_search_input()
        session.set_text(ti, "IT-2026-49625")
        session.assert_no_exceptions()
        dfs = session.get_dataframes()
        matched = [df for df in dfs if "Ticket ID" in df.columns and len(df) == 1]
        self.assertTrue(matched, "Expected 1-row dataframe matching search IT-2026-49625")
        self.assertEqual(matched[0].iloc[0]["Ticket ID"], "IT-2026-49625")
        session.set_text(session.find_search_input(), "")
        session.assert_no_exceptions()
        restored = [df for df in session.get_dataframes() if "Ticket ID" in df.columns and len(df) == 9]
        self.assertTrue(restored, "Clearing search should restore all 9 rows")

    def test_t1_f11_05_charts_and_dataframes_render_cleanly(self) -> None:
        session = AppSession().start()
        _upload_ref_csv_to_dashboard(session)
        charts = session.get_charts()
        self.assertGreaterEqual(len(charts), 1, "Expected at least one Altair/Vega-Lite chart on OMS Dashboard")
        for c in charts:
            self.assertTrue(c["spec"], "Vega-Lite chart spec should not be empty")
        dfs = session.get_dataframes()
        self.assertTrue(any(len(df) == 9 for df in dfs))


if __name__ == "__main__":
    unittest.main()
