"""Tier 2 E2E Tests: Boundary, Corner & Negative Cases (60 tests across F1-F11)."""

from __future__ import annotations

import csv
from datetime import date
import io
import unittest

import numpy as np
import pandas as pd

from tests_e2e.e2e_harness import (
    AppSession,
    VIEW_OMS_DASHBOARD,
    VIEW_OMS_FIR_COMPARE,
    assert_reference_converted_df,
    get_oms_converter,
    get_oms_exporters,
    load_xlsx_workbook_bytes,
    parse_csv_bytes,
)
from tests_e2e.fixtures import (
    ALT_TODAY,
    DATE_ONLY_RESOLVED_COL,
    DATETIME_COLS_15,
    DELTA_COLS_8,
    DISK_OMS_PATH,
    RAW_FORMULA_HINTS,
    RAW_RESOLVED_COL,
    REFERENCE_42_COLUMNS_DATE_ONLY,
    REFERENCE_42_COLUMNS_RAW,
    REFERENCE_TODAY,
    build_minimal_fir_df,
    build_reference_9_rows_df,
    corrupt_binary_bytes,
    to_csv_bytes,
    to_xls_html_bytes,
    to_xls_tsv_bytes,
    to_xlsx_bytes,
)


def _upload_df(session: AppSession, df: pd.DataFrame, filename: str = "oms.csv", set_today: date | None = REFERENCE_TODAY) -> AppSession:
    up = session.find_dashboard_uploader()
    session.upload_to(up, filename, to_csv_bytes(df), "text/csv")
    session.assert_no_exceptions()
    if set_today is not None:
        di = session.find_today_date_input()
        if di.value != set_today:
            session.set_date(di, set_today)
            session.assert_no_exceptions()
    return session


class TestTier2F1NavigationBoundaries(unittest.TestCase):
    def test_t2_f1_01_rapid_view_toggling_no_duplicate_keys(self) -> None:
        session = AppSession().start()
        for target in (
            VIEW_OMS_FIR_COMPARE,
            VIEW_OMS_DASHBOARD,
            VIEW_OMS_FIR_COMPARE,
            VIEW_OMS_DASHBOARD,
            VIEW_OMS_FIR_COMPARE,
        ):
            session.switch_view(target)
        self.assertEqual(session.get_nav_radio().value, VIEW_OMS_FIR_COMPARE)
        self.assertIn("EV Charger Complaint & Status Checker", session.extract_all_text())

    def test_t2_f1_02_invalid_nav_session_state_falls_back_gracefully(self) -> None:
        # 1. Assigning an out-of-options view to the navigation radio is rejected cleanly.
        s_invalid = AppSession().start()
        nav = s_invalid.get_nav_radio()
        with self.assertRaises(ValueError):
            nav.set_value("__NON_EXISTENT_VIEW__")
            s_invalid.rerun()

        # 2. Clearing/resetting nav & non-widget session state recovers cleanly to OMS Dashboard (index 0).
        session = AppSession().start()
        session.switch_view(VIEW_OMS_FIR_COMPARE)
        self.assertEqual(session.get_nav_radio().value, VIEW_OMS_FIR_COMPARE)

        nav_live = session.get_nav_radio()
        nav_key = getattr(nav_live, "key", None)
        session.set_session_state("__e2e_stale_nav_probe__", "stale_payload")
        session.rerun()
        del session.at.session_state["__e2e_stale_nav_probe__"]
        if nav_key is not None:
            session.set_session_state(nav_key, None)
        else:
            session.run_widget_action(lambda w=nav_live: w.set_value(VIEW_OMS_DASHBOARD))
        session.rerun()
        session.assert_no_exceptions()
        self.assertEqual(session.get_nav_radio().value, VIEW_OMS_DASHBOARD)
        _ = session.find_dashboard_uploader()

    def test_t2_f1_03_empty_datasets_both_views_render_cleanly(self) -> None:
        session = AppSession().start()
        session.assert_no_exceptions()
        _ = session.find_dashboard_uploader()
        session.switch_view(VIEW_OMS_FIR_COMPARE)
        session.assert_no_exceptions()
        self.assertIn("Step 1", session.extract_all_text())

    def test_t2_f1_04_active_dashboard_search_does_not_collide_with_compare(self) -> None:
        session = AppSession().start()
        _upload_df(session, build_reference_9_rows_df("iso"))
        session.set_text(session.find_search_input(), "EX-2026-32559")
        session.switch_view(VIEW_OMS_FIR_COMPARE)
        session.click_button("Sample Files")
        session.click_button("Compare Both Files")
        session.assert_no_exceptions()
        self.assertEqual(session.find_search_input().value, "")

    def test_t2_f1_05_oms_dashboard_always_index_zero(self) -> None:
        session = AppSession().start()
        nav = session.get_nav_radio()
        self.assertEqual(list(nav.options), [VIEW_OMS_DASHBOARD, VIEW_OMS_FIR_COMPARE])


class TestTier2F2IsolationBoundaries(unittest.TestCase):
    def test_t2_f2_01_corrupt_dashboard_upload_leaves_compare_intact(self) -> None:
        session = AppSession().start()
        session.switch_view(VIEW_OMS_FIR_COMPARE)
        session.click_button("Sample Files")
        session.click_button("Compare Both Files")
        session.switch_view(VIEW_OMS_DASHBOARD)
        up = session.find_dashboard_uploader()
        session.upload_to(up, "bad.xlsx", corrupt_binary_bytes(), "application/octet-stream")
        session.assert_no_exceptions()
        self.assertGreaterEqual(len(session.get_error_messages()), 1)
        session.switch_view(VIEW_OMS_FIR_COMPARE)
        self.assertIn("Charger & Ticket Check Results", session.extract_all_text())

    def test_t2_f2_02_corrupt_compare_upload_leaves_dashboard_intact(self) -> None:
        session = AppSession().start()
        _upload_df(session, build_reference_9_rows_df("iso"))
        session.switch_view(VIEW_OMS_FIR_COMPARE)
        fir_up, _ = session.find_compare_uploaders()
        session.upload_to(fir_up, "corrupt_fir.xlsx", corrupt_binary_bytes(), "application/octet-stream")
        session.assert_no_exceptions()
        self.assertGreaterEqual(len(session.get_error_messages()), 1)
        session.switch_view(VIEW_OMS_DASHBOARD)
        out_df = parse_csv_bytes(session.find_download(suffix=".csv")["bytes"])
        self.assertEqual(len(out_df), 9)

    def test_t2_f2_03_reupload_replaces_old_dashboard_rows_completely(self) -> None:
        session = AppSession().start()
        _upload_df(session, build_reference_9_rows_df("iso"), "first_9.csv")
        subset_3 = build_reference_9_rows_df("iso").iloc[:3].copy()
        _upload_df(session, subset_3, "second_3.csv")
        out_df = parse_csv_bytes(session.find_download(suffix=".csv")["bytes"])
        self.assertEqual(len(out_df), 3)
        self.assertEqual(set(out_df["Ticket ID"]), {"EX-2026-49803", "EX-2026-42245", "EX-2026-46397"})

    def test_t2_f2_04_compare_fir_only_stays_disabled_after_dashboard_oms_upload(self) -> None:
        session = AppSession().start()
        session.switch_view(VIEW_OMS_FIR_COMPARE)
        fir_up, _ = session.find_compare_uploaders()
        session.upload_to(fir_up, "fir.csv", to_csv_bytes(build_minimal_fir_df()), "text/csv")
        session.switch_view(VIEW_OMS_DASHBOARD)
        _upload_df(session, build_reference_9_rows_df("iso"))
        session.switch_view(VIEW_OMS_FIR_COMPARE)
        self.assertTrue(session.find_button("Compare Both Files").disabled)

    def test_t2_f2_05_clear_all_in_compare_preserves_dashboard_today_and_data(self) -> None:
        session = AppSession().start()
        _upload_df(session, build_reference_9_rows_df("iso"), set_today=ALT_TODAY)
        session.switch_view(VIEW_OMS_FIR_COMPARE)
        session.click_button("Sample Files")
        session.click_button("Clear All")
        session.switch_view(VIEW_OMS_DASHBOARD)
        self.assertEqual(session.find_today_date_input().value, ALT_TODAY)
        out_df = parse_csv_bytes(session.find_download(suffix=".csv")["bytes"])
        self.assertEqual(int(out_df.set_index("Ticket ID").loc["EX-2026-32559", "Hold total time(Days)"]), 25)


class TestTier2F3CompareBoundaries(unittest.TestCase):
    def test_t2_f3_01_empty_0_row_fir_handled_without_crash(self) -> None:
        session = AppSession().start()
        session.switch_view(VIEW_OMS_FIR_COMPARE)
        fir_up, _ = session.find_compare_uploaders()
        empty_fir = build_minimal_fir_df().iloc[0:0]
        session.upload_to(fir_up, "empty_fir.csv", to_csv_bytes(empty_fir), "text/csv")
        _, oms_up = session.find_compare_uploaders()
        session.upload_to(oms_up, "oms.csv", to_csv_bytes(build_reference_9_rows_df("iso")), "text/csv")
        session.assert_no_exceptions()
        compare_btn = session.find_button("Compare Both Files")
        self.assertFalse(compare_btn.disabled)
        session.click_button("Compare Both Files")
        session.assert_no_exceptions()

    def test_t2_f3_02_non_matching_search_in_compare_guards_zero_division(self) -> None:
        session = AppSession().start()
        session.switch_view(VIEW_OMS_FIR_COMPARE)
        session.click_button("Sample Files")
        session.click_button("Compare Both Files")
        session.set_text(session.find_search_input(), "ZZZZ_NO_MATCH_999")
        session.assert_no_exceptions()
        kpi = session.extract_kpi_text()
        self.assertIn("0.0%", kpi)

    def test_t2_f3_03_cascading_filter_and_reset_restores_all_controls(self) -> None:
        session = AppSession().start()
        session.switch_view(VIEW_OMS_FIR_COMPARE)
        session.click_button("Sample Files")
        session.click_button("Compare Both Files")
        session.set_multiselect(session.find_multiselect("Region / City"), ["East"])
        session.set_text(session.find_search_input(), "JH")
        session.click_button("Reset Filters")
        session.assert_no_exceptions()
        self.assertEqual(session.find_search_input().value, "")
        self.assertEqual(list(session.find_multiselect("Region / City").value), [])
        self.assertIn("All Chargers List (54)", session.get_tab_labels())

    def test_t2_f3_04_all_available_fir_zero_down_chargers_no_zero_division(self) -> None:
        session = AppSession().start()
        session.switch_view(VIEW_OMS_FIR_COMPARE)
        fir_up, _ = session.find_compare_uploaders()
        healthy_fir = build_minimal_fir_df(
            [
                {
                    "OCPP ID": "DL0022",
                    "Charger Status": "Available",
                    "Tenant": "ChargeZone - B2C",
                    "Priority": "3.Medium",
                    "Zone": "North",
                    "OEM": "Delta",
                    "Station Name": "Delhi Hub North",
                    "City": "Delhi",
                    "Issue Age": "0 Days",
                }
            ]
        )
        session.upload_to(fir_up, "healthy_fir.csv", to_csv_bytes(healthy_fir), "text/csv")
        _, oms_up = session.find_compare_uploaders()
        session.upload_to(oms_up, "oms.csv", to_csv_bytes(build_reference_9_rows_df("iso")), "text/csv")
        compare_btn = session.find_button("Compare Both Files")
        self.assertFalse(compare_btn.disabled)
        session.click_button("Compare Both Files")
        session.assert_no_exceptions()
        self.assertIn("0.0%", session.extract_kpi_text())

    def test_t2_f3_05_corrupt_fir_upload_shows_error_without_traceback(self) -> None:
        session = AppSession().start()
        session.switch_view(VIEW_OMS_FIR_COMPARE)
        fir_up, _ = session.find_compare_uploaders()
        session.upload_to(fir_up, "bad.xls", corrupt_binary_bytes(), "application/vnd.ms-excel")
        session.assert_no_exceptions()
        self.assertGreaterEqual(len(session.get_error_messages()), 1)


class TestTier2F4LoaderBoundaries(unittest.TestCase):
    def test_t2_f4_01_zero_byte_csv_caught_cleanly(self) -> None:
        session = AppSession().start()
        up = session.find_dashboard_uploader()
        session.upload_to(up, "empty.csv", b"", "text/csv")
        session.assert_no_exceptions()
        self.assertGreaterEqual(len(session.get_error_messages()), 1)

    def test_t2_f4_02_header_only_42_col_csv_loads_and_exports(self) -> None:
        header_only = pd.DataFrame(columns=list(REFERENCE_42_COLUMNS_RAW))
        session = AppSession().start()
        _upload_df(session, header_only, "header_only.csv", set_today=REFERENCE_TODAY)
        out_df = parse_csv_bytes(session.find_download(suffix=".csv")["bytes"])
        self.assertEqual(list(out_df.columns), list(REFERENCE_42_COLUMNS_DATE_ONLY))
        self.assertEqual(len(out_df), 0)

    def test_t2_f4_03_corrupt_xlsx_and_xls_raise_cleanly_and_show_st_error(self) -> None:
        conv = get_oms_converter()
        for fname in ("corrupt.xlsx", "corrupt.xls"):
            with self.assertRaises(Exception):
                conv.load_oms_raw_file(corrupt_binary_bytes(), fname)
        session = AppSession().start()
        up = session.find_dashboard_uploader()
        session.upload_to(up, "corrupt.xls", corrupt_binary_bytes(), "application/vnd.ms-excel")
        session.assert_no_exceptions()
        self.assertGreaterEqual(len(session.get_error_messages()), 1)

    def test_t2_f4_04_unsupported_extension_rejected_cleanly(self) -> None:
        conv = get_oms_converter()
        with self.assertRaises(ValueError):
            conv.load_oms_raw_file(b'{"a": 1}', "dump.json")

    def test_t2_f4_05_xlsx_with_pre_header_banner_rows_detected(self) -> None:
        raw_df = build_reference_9_rows_df("iso")
        xlsx_bytes = to_xlsx_bytes(raw_df, pre_header_rows=2)
        conv = get_oms_converter()
        loaded = conv.load_oms_raw_file(xlsx_bytes, "banner_oms.xlsx")
        out_df = conv.convert_oms_to_date_only(loaded, today_date=REFERENCE_TODAY)
        assert_reference_converted_df(out_df)

    def test_t2_f4_06_html_and_tsv_masquerading_as_xls_load_cleanly(self) -> None:
        raw_df = build_reference_9_rows_df("iso")
        conv = get_oms_converter()
        for fname, payload in (
            ("legacy_html.xls", to_xls_html_bytes(raw_df)),
            ("legacy_tsv.xls", to_xls_tsv_bytes(raw_df)),
        ):
            loaded = conv.load_oms_raw_file(payload, fname)
            out_df = conv.convert_oms_to_date_only(loaded, today_date=REFERENCE_TODAY)
            assert_reference_converted_df(out_df)


class TestTier2F5DateStripBoundaries(unittest.TestCase):
    def test_t2_f5_01_missing_and_sentinel_tokens_strip_to_empty(self) -> None:
        conv = get_oms_converter()
        for tok in (None, np.nan, pd.NaT, "", "   ", "-", "N/A", "NULL", "NaN", "na", "None"):
            self.assertEqual(conv.strip_to_date_str(tok), "", f"Failed for sentinel {tok!r}")

    def test_t2_f5_02_garbage_non_date_strings_strip_to_empty(self) -> None:
        conv = get_oms_converter()
        for bad in ("not-a-date", "colomn19", "today-colomn25", "TBD", "99/99/9999", "=U2-S2"):
            self.assertEqual(conv.strip_to_date_str(bad), "", f"Failed for garbage {bad!r}")

    def test_t2_f5_03_ambiguous_iso_vs_day_first_never_swaps_month_day(self) -> None:
        conv = get_oms_converter()
        self.assertEqual(conv.strip_to_date_str("2026-01-12"), "2026-01-12")
        self.assertEqual(conv.strip_to_date_str("2026-01-12 23:45:00"), "2026-01-12")
        self.assertEqual(conv.strip_to_date_str("12-01-2026"), "2026-01-12")
        self.assertEqual(conv.strip_to_date_str("12/01/2026 08:30"), "2026-01-12")

    def test_t2_f5_04_leap_day_and_year_boundary_dates(self) -> None:
        conv = get_oms_converter()
        self.assertEqual(conv.strip_to_date_str("2024-02-29 23:59:59"), "2024-02-29")
        self.assertEqual(conv.strip_to_date_str("2025-12-31 23:59:59"), "2025-12-31")
        self.assertEqual(conv.strip_to_date_str("2026-01-01 00:00:01"), "2026-01-01")

    def test_t2_f5_05_out_of_range_numeric_values_do_not_crash(self) -> None:
        conv = get_oms_converter()
        for num in (0, -5, 999999, float("inf"), float("-inf")):
            self.assertEqual(conv.strip_to_date_str(num), "")

    def test_t2_f5_06_microseconds_and_timezone_offsets_strip_cleanly(self) -> None:
        conv = get_oms_converter()
        self.assertEqual(conv.strip_to_date_str("2026-11-23 14:30:00.123456"), "2026-11-23")
        self.assertEqual(conv.strip_to_date_str("2026-11-23T00:05:00+05:30"), "2026-11-23")


class TestTier2F6RenameBoundaries(unittest.TestCase):
    def test_t2_f6_01_double_space_preserved_single_space_not_confused(self) -> None:
        conv = get_oms_converter()
        df = build_reference_9_rows_df("iso").iloc[[3]].copy()
        out_df = conv.convert_oms_to_date_only(df, today_date=REFERENCE_TODAY)
        self.assertIn(DATE_ONLY_RESOLVED_COL, out_df.columns)
        self.assertNotIn("Resolved total time(Days)", out_df.columns)

    def test_t2_f6_02_outer_whitespace_trimmed_preserving_internal_double_space(self) -> None:
        conv = get_oms_converter()
        df = build_reference_9_rows_df("iso").iloc[[3]].copy()
        xlsx_bytes = to_xlsx_bytes(df, header_overrides={RAW_RESOLVED_COL: f"  {RAW_RESOLVED_COL}  "})
        loaded = conv.load_oms_raw_file(xlsx_bytes, "padded_hdr.xlsx")
        out_df = conv.convert_oms_to_date_only(loaded, today_date=REFERENCE_TODAY)
        self.assertIn(DATE_ONLY_RESOLVED_COL, out_df.columns)
        self.assertEqual(int(out_df.iloc[0][DATE_ONLY_RESOLVED_COL]), 1)

    def test_t2_f6_03_missing_both_resolved_cols_no_key_error(self) -> None:
        conv = get_oms_converter()
        df = build_reference_9_rows_df("iso").drop(columns=[RAW_RESOLVED_COL, "Resolved At Time"])
        out_df = conv.convert_oms_to_date_only(df, today_date=REFERENCE_TODAY)
        self.assertNotIn(RAW_RESOLVED_COL, out_df.columns)
        self.assertEqual(len(out_df), 9)

    def test_t2_f6_04_overwrites_raw_days_hours_string_with_integer_delta(self) -> None:
        conv = get_oms_converter()
        df = build_reference_9_rows_df("iso").iloc[[3]].copy()
        df.loc[:, RAW_RESOLVED_COL] = "5 Days: 2 Hours"
        out_df = conv.convert_oms_to_date_only(df, today_date=REFERENCE_TODAY)
        self.assertEqual(int(out_df.iloc[0][DATE_ONLY_RESOLVED_COL]), 1)

    def test_t2_f6_05_overwrites_stale_numeric_value_when_resolved_at_blank(self) -> None:
        conv = get_oms_converter()
        df = build_reference_9_rows_df("iso").iloc[[0]].copy()
        df.loc[:, RAW_RESOLVED_COL] = "99"
        out_df = conv.convert_oms_to_date_only(df, today_date=REFERENCE_TODAY)
        self.assertTrue(pd.isna(out_df.iloc[0][DATE_ONLY_RESOLVED_COL]))


class TestTier2F7SchemaBoundaries(unittest.TestCase):
    def test_t2_f7_01_24_col_oms_dump_preserves_24_columns(self) -> None:
        self.assertTrue(DISK_OMS_PATH.exists())
        conv = get_oms_converter()
        raw_df = conv.load_oms_raw_file(DISK_OMS_PATH, "OMS DUMP.xlsx")
        out_df = conv.convert_oms_to_date_only(raw_df, today_date=REFERENCE_TODAY)
        self.assertEqual(len(out_df.columns), 24)
        self.assertEqual(list(out_df.columns), list(raw_df.columns))

    def test_t2_f7_02_hold_status_without_hold_delta_column_graceful(self) -> None:
        conv = get_oms_converter()
        df = build_reference_9_rows_df("iso").drop(columns=["Hold total time(Days)"])
        out_df = conv.convert_oms_to_date_only(df, today_date=REFERENCE_TODAY)
        self.assertEqual(len(out_df), 9)
        recalc = conv.recalculate_hold_days(out_df, today_date=ALT_TODAY)
        self.assertEqual(len(recalc), 9)

    def test_t2_f7_03_extra_43rd_custom_column_preserved(self) -> None:
        conv = get_oms_converter()
        df = build_reference_9_rows_df("iso")
        df["Custom Notes"] = [f"note-{i}" for i in range(len(df))]
        out_df = conv.convert_oms_to_date_only(df, today_date=REFERENCE_TODAY)
        self.assertEqual(len(out_df.columns), 43)
        self.assertEqual(out_df.columns[-1], "Custom Notes")
        self.assertEqual(list(out_df["Custom Notes"]), list(df["Custom Notes"]))

    def test_t2_f7_04_permuted_column_order_preserved(self) -> None:
        conv = get_oms_converter()
        df = build_reference_9_rows_df("iso")
        reversed_cols = list(reversed(list(df.columns)))
        df_rev = df[reversed_cols].copy()
        expected_cols = [DATE_ONLY_RESOLVED_COL if c == RAW_RESOLVED_COL else c for c in reversed_cols]
        out_df = conv.convert_oms_to_date_only(df_rev, today_date=REFERENCE_TODAY)
        self.assertEqual(list(out_df.columns), expected_cols)

    def test_t2_f7_05_unnamed_trailing_column_handled_cleanly(self) -> None:
        conv = get_oms_converter()
        df = build_reference_9_rows_df("iso")
        df["Unnamed: 42"] = np.nan
        out_df = conv.convert_oms_to_date_only(df, today_date=REFERENCE_TODAY)
        self.assertEqual(len(out_df), 9)
        self.assertIn("Unnamed: 42", out_df.columns)

    def test_t2_f7_06_subset_with_only_15_date_columns_no_delta_columns(self) -> None:
        conv = get_oms_converter()
        df = build_reference_9_rows_df("iso")[["Ticket ID", *DATETIME_COLS_15]].copy()
        out_df = conv.convert_oms_to_date_only(df, today_date=REFERENCE_TODAY)
        self.assertEqual(len(out_df), 9)
        self.assertEqual(str(out_df.loc[out_df["Ticket ID"] == "EX-2026-32559", "Assigned At"].iloc[0]), "2026-11-03")


class TestTier2F8DeltaBoundaries(unittest.TestCase):
    def test_t2_f8_01_row1_literal_formula_hints_ignored_and_blank_when_dates_absent(self) -> None:
        conv = get_oms_converter()
        df = build_reference_9_rows_df("iso").iloc[[0]].copy()
        for raw_col, hint in RAW_FORMULA_HINTS.items():
            self.assertEqual(df.iloc[0][raw_col], hint)
        out_df = conv.convert_oms_to_date_only(df, today_date=REFERENCE_TODAY)
        for dcol in DELTA_COLS_8:
            self.assertTrue(pd.isna(out_df.iloc[0][dcol]), f"{dcol} should be blank")

    def test_t2_f8_02_excel_equals_formula_strings_ignored(self) -> None:
        conv = get_oms_converter()
        df = build_reference_9_rows_df("iso").iloc[[0, 5]].copy()
        df.loc[:, "Acknowledgement time taken(Days)"] = "=U2-S2"
        df.loc[:, "Hold total time(Days)"] = "=TODAY()-Y2"
        out_df = conv.convert_oms_to_date_only(df, today_date=REFERENCE_TODAY)
        self.assertTrue(pd.isna(out_df.iloc[0]["Acknowledgement time taken(Days)"]))
        self.assertTrue(pd.isna(out_df.iloc[0]["Hold total time(Days)"]))
        self.assertEqual(int(out_df.iloc[1]["Acknowledgement time taken(Days)"]), 0)
        self.assertEqual(int(out_df.iloc[1]["Hold total time(Days)"]), 19)

    def test_t2_f8_03_zero_deltas_kept_as_integer_zero_not_blank(self) -> None:
        conv = get_oms_converter()
        out_df = conv.convert_oms_to_date_only(build_reference_9_rows_df("iso"), today_date=REFERENCE_TODAY)
        by_tid = out_df.set_index("Ticket ID")
        self.assertEqual(int(by_tid.loc["EX-2026-42245", "Acknowledgement time taken(Days)"]), 0)
        self.assertEqual(int(by_tid.loc["EX-2026-46397", "Hold total time(Days)"]), 0)
        self.assertEqual(int(by_tid.loc["EX-2026-34312", "Total time taken (Days)"]), 0)
        self.assertEqual(int(by_tid.loc["EX-2026-34312", "Approval Time taken (Days)"]), 0)

    def test_t2_f8_04_negative_deltas_signed_not_clamped(self) -> None:
        conv = get_oms_converter()
        df = build_reference_9_rows_df("iso").iloc[[5]].copy()
        out_df = conv.convert_oms_to_date_only(df, today_date=REFERENCE_TODAY)
        self.assertEqual(int(out_df.iloc[0]["Approval Time taken (Days)"]), -18)

        # Extreme year-boundary negative delta (-365)
        df2 = df.copy()
        df2.loc[:, "Approval Requested At Time"] = "2026-01-01 10:00:00"
        df2.loc[:, "Approval Accepted At Time"] = "2025-01-01 10:00:00"
        out_df2 = conv.convert_oms_to_date_only(df2, today_date=REFERENCE_TODAY)
        self.assertEqual(int(out_df2.iloc[0]["Approval Time taken (Days)"]), -365)

    def test_t2_f8_05_one_date_present_paired_date_missing_yields_blank(self) -> None:
        conv = get_oms_converter()
        df = build_reference_9_rows_df("iso").iloc[[0, 0]].copy().reset_index(drop=True)
        df.loc[0, "Assigned At"] = "2026-11-10 10:00:00"
        df.loc[0, "Closed At"] = ""
        df.loc[1, "Assigned At"] = ""
        df.loc[1, "Closed At"] = "2026-11-10 10:00:00"
        out_df = conv.convert_oms_to_date_only(df, today_date=REFERENCE_TODAY)
        self.assertTrue(pd.isna(out_df.iloc[0]["Total time taken (Days)"]))
        self.assertTrue(pd.isna(out_df.iloc[1]["Total time taken (Days)"]))

    def test_t2_f8_06_leap_year_vs_non_leap_year_day_diff(self) -> None:
        conv = get_oms_converter()
        df = build_reference_9_rows_df("iso").iloc[[0, 0]].copy().reset_index(drop=True)
        df.loc[0, "Assigned At"] = "2024-02-28 12:00:00"
        df.loc[0, "Acknowledged On"] = "2024-03-01 12:00:00"
        df.loc[1, "Assigned At"] = "2025-02-28 12:00:00"
        df.loc[1, "Acknowledged On"] = "2025-03-01 12:00:00"
        out_df = conv.convert_oms_to_date_only(df, today_date=REFERENCE_TODAY)
        self.assertEqual(int(out_df.iloc[0]["Acknowledgement time taken(Days)"]), 2)
        self.assertEqual(int(out_df.iloc[1]["Acknowledgement time taken(Days)"]), 1)


class TestTier2F9TodayBoundaries(unittest.TestCase):
    def test_t2_f9_01_today_earlier_than_hold_date_yields_negative_delta(self) -> None:
        conv = get_oms_converter()
        out_df = conv.convert_oms_to_date_only(build_reference_9_rows_df("iso"), today_date=date(2026, 11, 20))
        by_tid = out_df.set_index("Ticket ID")
        self.assertEqual(int(by_tid.loc["EX-2026-46397", "Hold total time(Days)"]), -5)

    def test_t2_f9_02_today_equals_hold_date_yields_zero(self) -> None:
        conv = get_oms_converter()
        out_df = conv.convert_oms_to_date_only(build_reference_9_rows_df("iso"), today_date=REFERENCE_TODAY)
        by_tid = out_df.set_index("Ticket ID")
        self.assertEqual(int(by_tid.loc["EX-2026-46397", "Hold total time(Days)"]), 0)

    def test_t2_f9_03_today_none_falls_back_to_system_today(self) -> None:
        conv = get_oms_converter()
        out_df = conv.convert_oms_to_date_only(build_reference_9_rows_df("iso"), today_date=None)
        expected = (date.today() - date(2026, 11, 6)).days
        by_tid = out_df.set_index("Ticket ID")
        self.assertEqual(int(by_tid.loc["EX-2026-32559", "Hold total time(Days)"]), expected)

    def test_t2_f9_04_today_changed_when_no_hold_dates_present(self) -> None:
        conv = get_oms_converter()
        df = build_reference_9_rows_df("iso")
        df["Hold Status Time"] = ""
        converted = conv.convert_oms_to_date_only(df, today_date=REFERENCE_TODAY)
        recalc = conv.recalculate_hold_days(converted, today_date=ALT_TODAY)
        self.assertTrue(recalc["Hold total time(Days)"].isna().all())

    def test_t2_f9_05_today_changed_across_leap_year_2024(self) -> None:
        conv = get_oms_converter()
        df = build_reference_9_rows_df("iso").iloc[[2]].copy()
        df.loc[:, "Hold Status Time"] = "2024-02-28 09:00:00"
        d0 = conv.convert_oms_to_date_only(df, today_date=date(2024, 2, 28))
        d2 = conv.recalculate_hold_days(d0, today_date=date(2024, 3, 1))
        self.assertEqual(int(d0.iloc[0]["Hold total time(Days)"]), 0)
        self.assertEqual(int(d2.iloc[0]["Hold total time(Days)"]), 2)


class TestTier2F10ExportBoundaries(unittest.TestCase):
    def test_t2_f10_01_export_zero_row_dataframe_csv_and_xlsx(self) -> None:
        conv = get_oms_converter()
        export_csv, export_xlsx = get_oms_exporters()
        empty_raw = pd.DataFrame(columns=list(REFERENCE_42_COLUMNS_RAW))
        converted = conv.convert_oms_to_date_only(empty_raw, today_date=REFERENCE_TODAY)
        csv_data = export_csv(converted)
        csv_bytes = csv_data.encode("utf-8") if isinstance(csv_data, str) else csv_data
        parsed = parse_csv_bytes(csv_bytes)
        self.assertEqual(list(parsed.columns), list(REFERENCE_42_COLUMNS_DATE_ONLY))
        self.assertEqual(len(parsed), 0)

        wb = load_xlsx_workbook_bytes(export_xlsx(converted))
        ws = wb.active
        self.assertEqual(ws.max_row, 1)
        self.assertEqual([ws.cell(1, c).value for c in range(1, ws.max_column + 1)], list(REFERENCE_42_COLUMNS_DATE_ONLY))

    def test_t2_f10_02_export_csv_mixed_zero_negative_na_no_float_suffix(self) -> None:
        conv = get_oms_converter()
        export_csv, _ = get_oms_exporters()
        converted = conv.convert_oms_to_date_only(build_reference_9_rows_df("iso"), today_date=REFERENCE_TODAY)
        csv_data = export_csv(converted)
        text = csv_data.decode("utf-8") if isinstance(csv_data, bytes) else csv_data
        rows = {r["Ticket ID"]: r for r in csv.DictReader(io.StringIO(text))}
        self.assertEqual(rows["EX-2026-42245"]["Acknowledgement time taken(Days)"], "0")
        self.assertEqual(rows["EX-2026-32559"]["Approval Time taken (Days)"], "-18")
        self.assertEqual(rows["EX-2026-32559"]["Hold total time(Days)"], "19")
        self.assertEqual(rows["EX-2026-49803"]["Hold total time(Days)"], "")
        self.assertNotIn(".0", text)

    def test_t2_f10_03_export_xlsx_preserves_int_types_and_none_blanks(self) -> None:
        conv = get_oms_converter()
        _, export_xlsx = get_oms_exporters()
        converted = conv.convert_oms_to_date_only(build_reference_9_rows_df("iso"), today_date=REFERENCE_TODAY)
        wb = load_xlsx_workbook_bytes(export_xlsx(converted))
        ws = wb.active
        headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
        app_c = headers.index("Approval Time taken (Days)") + 1
        tid_c = headers.index("Ticket ID") + 1
        vals = {ws.cell(r, tid_c).value: ws.cell(r, app_c).value for r in range(2, ws.max_row + 1)}
        self.assertEqual(vals["EX-2026-32559"], -18)
        self.assertIs(type(vals["EX-2026-32559"]), int)
        self.assertEqual(vals["EX-2026-34312"], 0)
        self.assertIs(type(vals["EX-2026-34312"]), int)
        self.assertIsNone(vals["EX-2026-49803"])

    def test_t2_f10_04_round_trip_csv_and_xlsx_exports_reingest_identically(self) -> None:
        conv = get_oms_converter()
        export_csv, export_xlsx = get_oms_exporters()
        converted1 = conv.convert_oms_to_date_only(build_reference_9_rows_df("iso"), today_date=REFERENCE_TODAY)
        csv_data = export_csv(converted1)
        csv_bytes = csv_data.encode("utf-8") if isinstance(csv_data, str) else csv_data
        xlsx_bytes = export_xlsx(converted1)

        re_csv = conv.convert_oms_to_date_only(conv.load_oms_raw_file(csv_bytes, "re.csv"), today_date=REFERENCE_TODAY)
        re_xlsx = conv.convert_oms_to_date_only(conv.load_oms_raw_file(xlsx_bytes, "re.xlsx"), today_date=REFERENCE_TODAY)
        assert_reference_converted_df(re_csv)
        assert_reference_converted_df(re_xlsx)

    def test_t2_f10_05_special_characters_commas_quotes_pipes_unicode_preserved(self) -> None:
        conv = get_oms_converter()
        export_csv, export_xlsx = get_oms_exporters()
        df = build_reference_9_rows_df("iso")
        special = 'JK | Srinagar, "Arise" ⚡ №1'
        df.loc[df.index[0], "Station"] = special
        converted = conv.convert_oms_to_date_only(df, today_date=REFERENCE_TODAY)

        csv_data = export_csv(converted)
        csv_bytes = csv_data.encode("utf-8") if isinstance(csv_data, str) else csv_data
        parsed_csv = parse_csv_bytes(csv_bytes)
        self.assertEqual(len(parsed_csv.columns), 42)
        self.assertEqual(parsed_csv.iloc[0]["Station"], special)

        wb = load_xlsx_workbook_bytes(export_xlsx(converted))
        ws = wb.active
        headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
        st_c = headers.index("Station") + 1
        self.assertEqual(ws.cell(2, st_c).value, special)

    def test_t2_f10_06_csv_export_bytes_and_xlsx_headers_consistent_after_hold_recalc(self) -> None:
        conv = get_oms_converter()
        export_csv, export_xlsx = get_oms_exporters()
        c1 = conv.convert_oms_to_date_only(build_reference_9_rows_df("iso"), today_date=REFERENCE_TODAY)
        c2 = conv.recalculate_hold_days(c1, today_date=ALT_TODAY)
        csv_data = export_csv(c2)
        csv_bytes = csv_data.encode("utf-8") if isinstance(csv_data, str) else csv_data
        df_csv = parse_csv_bytes(csv_bytes)
        df_xlsx = conv.load_oms_raw_file(export_xlsx(c2), "recalc.xlsx")
        self.assertEqual(list(df_csv.columns), list(df_xlsx.columns))
        self.assertEqual(int(df_csv.set_index("Ticket ID").loc["EX-2026-40520", "Hold total time(Days)"]), 16)
        self.assertEqual(int(df_xlsx.set_index("Ticket ID").loc["EX-2026-40520", "Hold total time(Days)"]), 16)


class TestTier2F11DashboardUIBoundaries(unittest.TestCase):
    def test_t2_f11_01_empty_filter_match_renders_zero_gracefully(self) -> None:
        session = AppSession().start()
        _upload_df(session, build_reference_9_rows_df("iso"))
        session.set_text(session.find_search_input(), "NONEXISTENT_TICKET_XYZ_999")
        session.assert_no_exceptions()
        self.assertRegex(session.extract_kpi_text(), r"\b0\b")

    def test_t2_f11_02_regex_special_chars_in_live_search_no_re_error(self) -> None:
        session = AppSession().start()
        df = build_reference_9_rows_df("iso")
        df.loc[df.index[0], "Station"] = "[O&M] | (urgent)* \\ check"
        _upload_df(session, df)
        session.set_text(session.find_search_input(), "[O&M]|(*\\")
        session.assert_no_exceptions()
        session.set_text(session.find_search_input(), "[O&M]")
        session.assert_no_exceptions()
        dfs = [d for d in session.get_dataframes() if "Ticket ID" in d.columns and len(d) == 1]
        self.assertTrue(dfs, "Expected literal substring match for '[O&M]'")

    def test_t2_f11_03_null_metadata_fields_do_not_crash_filters_or_charts(self) -> None:
        session = AppSession().start()
        df = build_reference_9_rows_df("iso")
        for col in ("Station", "Location", "Charger", "Zone", "Assigned To"):
            df.loc[df.index[0], col] = np.nan
        _upload_df(session, df)
        session.assert_no_exceptions()
        self.assertGreaterEqual(len(session.get_charts()), 1)

    def test_t2_f11_04_large_2605_row_oms_dump_renders_within_timeout(self) -> None:
        self.assertTrue(DISK_OMS_PATH.exists())
        session = AppSession().start()
        up = session.find_dashboard_uploader()
        session.upload_to(
            up,
            "OMS DUMP.xlsx",
            DISK_OMS_PATH.read_bytes(),
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        session.assert_no_exceptions()
        kpi = session.extract_kpi_text()
        self.assertTrue("2,605" in kpi or "2605" in kpi)
        self.assertGreaterEqual(len(session.get_charts()), 1)

    def test_t2_f11_05_all_closed_zero_breach_dataset_no_zero_division(self) -> None:
        session = AppSession().start()
        df = build_reference_9_rows_df("iso").iloc[[6]].copy()  # Closed, No breach, no Hold
        _upload_df(session, df)
        session.assert_no_exceptions()
        kpi = session.extract_kpi_text()
        self.assertNotIn("nan", kpi.lower())


if __name__ == "__main__":
    unittest.main()
