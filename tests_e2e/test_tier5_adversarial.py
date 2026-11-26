"""Tier 5 E2E Tests: White-Box Adversarial Hardening (Findings 1-8 & R2-1..R2-4)."""

from __future__ import annotations

import csv
from datetime import date, datetime, timezone
import io
import struct
import unittest
from unittest.mock import patch
import warnings

import openpyxl
import pandas as pd

from tests_e2e.e2e_harness import (
    AppSession,
    VIEW_OMS_DASHBOARD,
    VIEW_OMS_FIR_COMPARE,
    get_oms_converter,
    get_oms_exporters,
    load_xlsx_workbook_bytes,
    parse_csv_bytes,
)
from tests_e2e.fixtures import (
    REFERENCE_TODAY,
    build_minimal_fir_df,
    build_reference_9_rows_df,
    to_csv_bytes,
    to_xlsx_bytes,
)


def _biff_record(rec_id: int, payload: bytes) -> bytes:
    return struct.pack("<HH", rec_id, len(payload)) + payload


def _biff_short_ascii(text: str) -> bytes:
    raw = text.encode("latin-1")
    return struct.pack("<BB", len(raw), 0x00) + raw


def _biff_unicode_str_ascii(text: str) -> bytes:
    raw = text.encode("latin-1")
    return struct.pack("<HB", len(raw), 0x00) + raw


def _wrap_ole2_v3(wb_stream: bytes) -> bytes:
    if len(wb_stream) < 4096:
        wb_stream += b"\x00" * (4096 - len(wb_stream))

    sec_size = 512
    wb_sectors = (len(wb_stream) + sec_size - 1) // sec_size
    wb_padded = wb_stream.ljust(wb_sectors * sec_size, b"\x00")

    fat_entries = [0xFFFFFFFF] * max(128, ((2 + wb_sectors + 127) // 128) * 128)
    fat_entries[0] = 0xFFFFFFFD
    fat_entries[1] = 0xFFFFFFFE
    for i in range(wb_sectors):
        sec = 2 + i
        fat_entries[sec] = (sec + 1) if i + 1 < wb_sectors else 0xFFFFFFFE
    fat_sector = struct.pack("<128I", *fat_entries[:128])

    def _dir_entry(name: str, obj_type: int, start_sec: int, size: int) -> bytes:
        name_u16 = (name + "\x00").encode("utf-16-le")
        buf = bytearray(128)
        buf[: len(name_u16)] = name_u16
        struct.pack_into("<H", buf, 0x40, len(name_u16))
        buf[0x42] = obj_type
        buf[0x43] = 1
        struct.pack_into("<III", buf, 0x44, 0xFFFFFFFF, 0xFFFFFFFF, 0xFFFFFFFF)
        if obj_type == 5:
            struct.pack_into("<I", buf, 0x4C, 1)
        struct.pack_into("<II", buf, 0x74, start_sec, size)
        return bytes(buf)

    dir_sector = (
        _dir_entry("Root Entry", 5, 0xFFFFFFFE, 0)
        + _dir_entry("Workbook", 2, 2, len(wb_stream))
        + (b"\x00" * 256)
    )

    hdr = bytearray(512)
    hdr[:8] = b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1"
    struct.pack_into("<HHHH", hdr, 0x18, 0x003E, 0x0003, 0xFFFE, 9)
    struct.pack_into("<H", hdr, 0x20, 6)
    struct.pack_into("<IIIIIII", hdr, 0x2C, 1, 1, 0, 4096, 0xFFFFFFFE, 0, 0xFFFFFFFE)
    struct.pack_into("<I", hdr, 0x48, 0)
    difat = [0] + [0xFFFFFFFF] * 108
    struct.pack_into("<109I", hdr, 0x4C, *difat)

    return bytes(hdr) + fat_sector + dir_sector + wb_padded


def _build_biff8_with_nested_chart() -> bytes:
    """Build a BIFF8 .xls whose worksheet embeds a nested Chart BOF/EOF before cell records."""
    bof_global = _biff_record(0x0809, struct.pack("<HHHHII", 0x0600, 0x0005, 0, 0, 0, 0))
    codepage = _biff_record(0x0042, struct.pack("<H", 1252))
    xf0 = struct.pack("<HHH", 0, 0, 0x0001) + b"\x00" * 14
    xf_rec = _biff_record(0x00E0, xf0)

    strings = ["Ticket ID", "Status", "EX-NEST-1", "Open"]
    sst_payload = bytearray(struct.pack("<II", len(strings), len(strings)))
    for s in strings:
        sst_payload.extend(_biff_unicode_str_ascii(s))
    sst_rec = _biff_record(0x00FC, bytes(sst_payload))
    eof_rec = _biff_record(0x000A, b"")

    sname = _biff_short_ascii("Sheet1")
    sheet_offset = (
        len(bof_global)
        + len(codepage)
        + len(xf_rec)
        + (4 + 6 + len(sname))
        + len(sst_rec)
        + len(eof_rec)
    )
    boundsheet = _biff_record(0x0085, struct.pack("<IH", sheet_offset, 0) + sname)

    # Worksheet BOF -> embedded Chart BOF (0x0020) + bogus cell + Chart EOF -> real cell records -> Worksheet EOF
    ws = bytearray(_biff_record(0x0809, struct.pack("<HHHHII", 0x0600, 0x0010, 0, 0, 0, 0)))
    ws.extend(_biff_record(0x0809, struct.pack("<HHHHII", 0x0600, 0x0020, 0, 0, 0, 0)))
    ws.extend(_biff_record(0x0204, struct.pack("<HHH", 0, 0, 0) + _biff_unicode_str_ascii("CHART_NOISE")))
    ws.extend(eof_rec)  # closes nested Chart substream, NOT worksheet
    ws.extend(_biff_record(0x00FD, struct.pack("<HHHI", 0, 0, 0, 0)))
    ws.extend(_biff_record(0x00FD, struct.pack("<HHHI", 0, 1, 0, 1)))
    ws.extend(_biff_record(0x00FD, struct.pack("<HHHI", 1, 0, 0, 2)))
    ws.extend(_biff_record(0x00FD, struct.pack("<HHHI", 1, 1, 0, 3)))
    ws.extend(eof_rec)

    wb_stream = bof_global + codepage + xf_rec + boundsheet + sst_rec + eof_rec + bytes(ws)
    return _wrap_ole2_v3(wb_stream)


class TestTier5Adversarial(unittest.TestCase):
    def test_t5_01_excel_export_sanitizes_control_characters_in_dashboard_and_compare(self) -> None:
        from core.exporter import generate_reconciliation_excel
        from core.reconciler import reconcile_incident_records

        _, export_xlsx = get_oms_exporters()

        # 1. Core OMS date_only Excel sanitization
        ctrl_df = pd.DataFrame(
            [
                {
                    "Ticket ID": "EX-CTRL-1",
                    "Status": "Open",
                    "Remarks": "Bad\x0bChar\x0cHere\x00End",
                    "Station": "S\x071\nLine2",
                }
            ]
        )
        xlsx_bytes = export_xlsx(ctrl_df)
        wb = load_xlsx_workbook_bytes(xlsx_bytes)
        ws = wb.active
        self.assertEqual(ws.cell(2, 3).value, "BadCharHereEnd")
        self.assertEqual(ws.cell(2, 4).value, "S1\nLine2")

        # 2. Core Compare reconciliation Excel sanitization
        fir_ctrl = build_minimal_fir_df(
            [
                {
                    "OCPP ID": "CHG-CTRL-1",
                    "Status": "Faulted",
                    "Error Info": "Bad\x0bFault\x0cCode",
                    "Station Name": "St\x01A",
                }
            ]
        )
        oms_min = pd.DataFrame([{"Ticket ID": "T-1", "Charger": "CHG-OTHER", "Status": "Open"}])
        rec = reconcile_incident_records(fir_ctrl, oms_min)
        rec_bytes = generate_reconciliation_excel(rec)
        rec_wb = load_xlsx_workbook_bytes(rec_bytes)
        self.assertIn("Ticket Generation Gaps", rec_wb.sheetnames)
        gap_ws = rec_wb["Ticket Generation Gaps"]
        gap_headers = [gap_ws.cell(1, c).value for c in range(1, gap_ws.max_column + 1)]
        err_col = gap_headers.index("Error Info") + 1
        self.assertEqual(gap_ws.cell(2, err_col).value, "BadFaultCode")

        # 3. UI OMS Dashboard upload with control characters
        session = AppSession().start()
        csv_payload = "Ticket ID,Status,Remarks\nEX-1,Open,Bad\x0bChar\n".encode("utf-8")
        session.upload_to(session.find_dashboard_uploader(), "ctrl.csv", csv_payload, "text/csv")
        session.assert_no_exceptions()
        dl_wb = load_xlsx_workbook_bytes(session.find_download(suffix=".xlsx")["bytes"])
        self.assertEqual(dl_wb.active.cell(2, 3).value, "BadChar")

        # 4. UI OMS-FIR Compare upload with control characters
        session.switch_view(VIEW_OMS_FIR_COMPARE)
        fir_up, _ = session.find_compare_uploaders()
        session.upload_to(fir_up, "fir_ctrl.csv", to_csv_bytes(fir_ctrl), "text/csv")
        _, oms_up = session.find_compare_uploaders()
        session.upload_to(oms_up, "oms_min.csv", to_csv_bytes(oms_min), "text/csv")
        session.click_button("Compare Both Files")
        session.assert_no_exceptions()
        cmp_wb = load_xlsx_workbook_bytes(session.find_download(suffix=".xlsx")["bytes"])
        self.assertIn("Executive Summary", cmp_wb.sheetnames)

    def test_t5_02_compare_view_nonstandard_status_col_and_mixed_type_zone_no_crash(self) -> None:
        # 1. FIR with FIR_Status (no literal "Status" column) + Charger Make
        fir_alt = pd.DataFrame(
            [
                {
                    "OCPP ID": "CHG-ALT-1",
                    "FIR_Status": "Faulted",
                    "Tenant Name": "ClientA",
                    "Priority": "Critical",
                    "Zone": "West",
                    "Charger Make": "Delta",
                    "Station Name": "West Hub",
                    "Issue_Age": "2 Days",
                }
            ]
        )
        oms_min = pd.DataFrame([{"Ticket ID": "T-ALT-1", "Charger": "CHG-ALT-1", "Status": "Open"}])

        s1 = AppSession().start()
        s1.switch_view(VIEW_OMS_FIR_COMPARE)
        fir_up, _ = s1.find_compare_uploaders()
        s1.upload_to(fir_up, "fir_alt.csv", to_csv_bytes(fir_alt), "text/csv")
        _, oms_up = s1.find_compare_uploaders()
        s1.upload_to(oms_up, "oms_min.csv", to_csv_bytes(oms_min), "text/csv")
        s1.click_button("Compare Both Files")
        s1.assert_no_exceptions()
        self.assertIn("No brand data matching current filters.", s1.extract_all_text())

        # 2. FIR .xlsx with mixed str/int values in Zone, Tenant Name, Charger Make
        fir_mixed = pd.DataFrame(
            [
                {
                    "OCPP ID": "CHG-MIX-1",
                    "Status": "Faulted",
                    "Tenant Name": "ClientA",
                    "Priority": "Critical",
                    "Zone": "West",
                    "Charger Make": "Delta",
                    "Station Name": "West Hub",
                    "Issue_Age": "2 Days",
                },
                {
                    "OCPP ID": "CHG-MIX-2",
                    "Status": "Available",
                    "Tenant Name": 202,
                    "Priority": "Low",
                    "Zone": 101,
                    "Charger Make": 303,
                    "Station Name": "Numeric Zone Hub",
                    "Issue_Age": "0 Days",
                },
            ]
        )
        s2 = AppSession().start()
        s2.switch_view(VIEW_OMS_FIR_COMPARE)
        fir_up2, _ = s2.find_compare_uploaders()
        s2.upload_to(
            fir_up2,
            "fir_mixed.xlsx",
            to_xlsx_bytes(fir_mixed),
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        _, oms_up2 = s2.find_compare_uploaders()
        s2.upload_to(oms_up2, "oms_min.csv", to_csv_bytes(oms_min), "text/csv")
        s2.click_button("Compare Both Files")
        s2.assert_no_exceptions()

        zone_ms = s2.find_multiselect("Region / City")
        self.assertEqual(list(zone_ms.options), ["101", "West"])
        s2.set_multiselect(zone_ms, ["101"])
        s2.assert_no_exceptions()
        self.assertIn("Showing 1 of 2 chargers", s2.extract_all_text())

    def test_t5_03_duplicate_headers_after_whitespace_strip_do_not_crash(self) -> None:
        conv = get_oms_converter()
        csv_bytes = (
            b"Ticket ID,Created At,Created At ,Status\n"
            b"EX-1,2026-11-20 10:00:00,2026-11-21 10:00:00,Open\n"
        )
        loaded_csv = conv.load_oms_raw_file(csv_bytes, "dup_space.csv")
        self.assertEqual(list(loaded_csv.columns), ["Ticket ID", "Created At", "Created At.1", "Status"])
        conv_csv = conv.convert_oms_to_date_only(loaded_csv, today_date=REFERENCE_TODAY)
        self.assertEqual(conv_csv.loc[0, "Created At"], "2026-11-20")

        # Also test XLSX with duplicate headers differing only by trailing whitespace
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["Ticket ID", "Created At", "Created At ", "Status"])
        ws.append(["EX-2", "2026-11-22 08:30:00", "2026-11-23 09:30:00", "Closed"])
        buf = io.BytesIO()
        wb.save(buf)
        loaded_xlsx = conv.load_oms_raw_file(buf.getvalue(), "dup_space.xlsx")
        self.assertEqual(list(loaded_xlsx.columns), ["Ticket ID", "Created At", "Created At.1", "Status"])
        conv_xlsx = conv.convert_oms_to_date_only(loaded_xlsx, today_date=REFERENCE_TODAY)
        self.assertEqual(conv_xlsx.loc[0, "Created At"], "2026-11-22")

        # UI upload verification
        session = AppSession().start()
        session.upload_to(session.find_dashboard_uploader(), "dup_space.csv", csv_bytes, "text/csv")
        session.assert_no_exceptions()
        self.assertEqual(session.get_error_messages(), [])
        out_df = parse_csv_bytes(session.find_download(suffix=".csv")["bytes"])
        self.assertEqual(len(out_df), 1)
        self.assertEqual(out_df.loc[0, "Created At"], "2026-11-20")

    def test_t5_04_row1_banner_with_single_known_header_name_falls_back_to_real_header(self) -> None:
        conv = get_oms_converter()

        # 1. CSV banner starting with "Updated At" above real 5-col OMS header
        csv_banner = (
            b"Updated At,2026-11-25 10:00:00\n"
            b"Sno,Ticket ID,Assigned At,Acknowledged On,Status\n"
            b"1,EX-BAN-1,2026-11-20 09:00:00,2026-11-22 10:00:00,Open\n"
        )
        df_csv = conv.load_oms_raw_file(csv_banner, "banner_updated_at.csv")
        self.assertEqual(list(df_csv.columns), ["Sno", "Ticket ID", "Assigned At", "Acknowledged On", "Status"])
        self.assertEqual(df_csv.shape, (1, 5))
        conv_csv = conv.convert_oms_to_date_only(df_csv, today_date=REFERENCE_TODAY)
        self.assertEqual(int(conv_csv.loc[0, "Acknowledgement time taken(Days)"]), 2)

        # 2. XLSX banner containing single cell "Ticket ID" above real header
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["Ticket ID"])
        ws.append(["Sno", "Ticket ID", "Assigned At", "Acknowledged On", "Status"])
        ws.append([1, "EX-BAN-2", "2026-11-21 09:00:00", "2026-11-24 10:00:00", "Open"])
        buf = io.BytesIO()
        wb.save(buf)
        df_xlsx = conv.load_oms_raw_file(buf.getvalue(), "banner_ticket_id.xlsx")
        self.assertEqual(list(df_xlsx.columns), ["Sno", "Ticket ID", "Assigned At", "Acknowledged On", "Status"])
        self.assertEqual(df_xlsx.shape, (1, 5))

        # 3. Single-column dump still loads cleanly on row 1
        single_col_csv = b"Ticket ID\nEX-1\nEX-2\n"
        df_single = conv.load_oms_raw_file(single_col_csv, "single.csv")
        self.assertEqual(list(df_single.columns), ["Ticket ID"])
        self.assertEqual(len(df_single), 2)

    def test_t5_05_excel_export_writes_equals_strings_as_text_not_formulas(self) -> None:
        from core.exporter import generate_reconciliation_excel
        from core.reconciler import reconcile_incident_records

        conv = get_oms_converter()
        export_csv, export_xlsx = get_oms_exporters()

        raw_df = build_reference_9_rows_df("iso").iloc[[5]].copy()  # Approval delta = -18, Ack delta = 0
        raw_df.loc[raw_df.index[0], "Station"] = "=1+1"
        raw_df.loc[raw_df.index[0], "Remarks"] = "=CMD|'/C calc'!A0"
        conv_df = conv.convert_oms_to_date_only(raw_df, today_date=REFERENCE_TODAY)

        xlsx_bytes = export_xlsx(conv_df)
        for data_only in (False, True):
            wb = openpyxl.load_workbook(io.BytesIO(xlsx_bytes), data_only=data_only)
            ws = wb.active
            headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
            st_cell = ws.cell(2, headers.index("Station") + 1)
            rm_cell = ws.cell(2, headers.index("Remarks") + 1)
            app_cell = ws.cell(2, headers.index("Approval Time taken (Days)") + 1)
            ack_cell = ws.cell(2, headers.index("Acknowledgement time taken(Days)") + 1)

            self.assertEqual(st_cell.value, "=1+1")
            self.assertEqual(st_cell.data_type, "s")
            self.assertEqual(rm_cell.value, "=CMD|'/C calc'!A0")
            self.assertEqual(rm_cell.data_type, "s")
            self.assertEqual(app_cell.value, -18)
            self.assertIs(type(app_cell.value), int)
            self.assertEqual(ack_cell.value, 0)
            self.assertIs(type(ack_cell.value), int)

        # CSV export keeps literal text and negative integers intact
        csv_text = export_csv(conv_df)
        csv_rows = list(csv.DictReader(io.StringIO(csv_text)))
        self.assertEqual(csv_rows[0]["Station"], "=1+1")
        self.assertEqual(csv_rows[0]["Approval Time taken (Days)"], "-18")

        # Compare reconciliation Excel export also stores '='-prefixed text as string
        fir_eq = build_minimal_fir_df(
            [{"OCPP ID": "CHG-EQ-1", "Status": "Faulted", "Station Name": "=SUM(A1:A2)"}]
        )
        oms_eq = pd.DataFrame([{"Ticket ID": "T-1", "Charger": "CHG-OTHER", "Status": "Open"}])
        rec_bytes = generate_reconciliation_excel(reconcile_incident_records(fir_eq, oms_eq))
        rec_wb = openpyxl.load_workbook(io.BytesIO(rec_bytes), data_only=False)
        gap_ws = rec_wb["Ticket Generation Gaps"]
        gap_headers = [gap_ws.cell(1, c).value for c in range(1, gap_ws.max_column + 1)]
        station_cell = gap_ws.cell(2, gap_headers.index("Station Name") + 1)
        self.assertEqual(station_cell.value, "=SUM(A1:A2)")
        self.assertEqual(station_cell.data_type, "s")

    def test_t5_06_biff8_nested_chart_bof_eof_and_html4_unclosed_tags(self) -> None:
        conv = get_oms_converter()

        # 1. BIFF8 with nested Chart BOF/EOF before worksheet data
        biff_nested = _build_biff8_with_nested_chart()
        df_biff = conv.load_oms_raw_file(biff_nested, "nested_chart.xls")
        self.assertEqual(list(df_biff.columns), ["Ticket ID", "Status"])
        self.assertEqual(len(df_biff), 1)
        self.assertEqual(df_biff.loc[0, "Ticket ID"], "EX-NEST-1")
        self.assertEqual(df_biff.loc[0, "Status"], "Open")

        # 2. Legacy HTML4 table with unclosed <tr>, <th>, <td> tags (and even unclosed </table>)
        html4_unclosed = (
            b"<html><body><table>"
            b"<tr><th>Ticket ID<th>Status"
            b"<tr><td>EX-HTML4-1<td>Open"
            b"<tr><td>EX-HTML4-2<td>Closed"
            b"</table></body></html>"
        )
        df_html = conv.load_oms_raw_file(html4_unclosed, "html4_unclosed.xls")
        self.assertEqual(list(df_html.columns), ["Ticket ID", "Status"])
        self.assertEqual(df_html.to_dict(orient="records"), [
            {"Ticket ID": "EX-HTML4-1", "Status": "Open"},
            {"Ticket ID": "EX-HTML4-2", "Status": "Closed"},
        ])

        html4_no_close_table = b"<table><tr><th>Ticket ID<th>Status<tr><td>EX-EOF-1<td>On Hold"
        df_html_eof = conv.load_oms_raw_file(html4_no_close_table, "html4_eof.xls")
        self.assertEqual(list(df_html_eof.columns), ["Ticket ID", "Status"])
        self.assertEqual(df_html_eof.loc[0, "Ticket ID"], "EX-EOF-1")

    def test_t5_07_date_parsing_dot_formats_and_year_bounds(self) -> None:
        conv = get_oms_converter()
        with warnings.catch_warnings(record=True) as caught:
            warnings.simplefilter("always")

            # Dot-separated and 2-digit-year %d-%b-%y %H:%M formats
            self.assertEqual(conv.parse_date_only("25.11.2026"), date(2026, 11, 25))
            self.assertEqual(conv.parse_date_only("25.11.2026 14:30"), date(2026, 11, 25))
            self.assertEqual(conv.parse_date_only("2026.11.25"), date(2026, 11, 25))
            self.assertEqual(conv.parse_date_only("2026.11.25 08:15:30"), date(2026, 11, 25))
            self.assertEqual(conv.parse_date_only("25-Nov-26 14:30"), date(2026, 11, 25))
            self.assertEqual(conv.parse_date_only("25-Nov-26 02:30 PM"), date(2026, 11, 25))

            # Consistent [1970, 2100] year bounds across datetime, date, Timestamp, and Excel serials
            for out_of_bounds in (
                datetime(1969, 12, 31, 23, 59),
                date(1969, 12, 31),
                pd.Timestamp("1969-12-31"),
                datetime(2101, 1, 1, 0, 0),
                date(9999, 12, 31),
                20000,
                "20000",
                80000,
                "80000",
                "10.5",
                "2026.11",
                "1.2.3",
                "=A1-B1",
                "colomn21-colomn19",
            ):
                self.assertIsNone(conv.parse_date_only(out_of_bounds), f"Expected None for {out_of_bounds!r}")
                self.assertEqual(conv.strip_to_date_str(out_of_bounds), "", f"Expected '' for {out_of_bounds!r}")

            self.assertEqual(conv.parse_date_only(datetime(1970, 1, 1, 8, 0)), date(1970, 1, 1))
            self.assertEqual(conv.parse_date_only(date(2100, 12, 31)), date(2100, 12, 31))
            self.assertEqual(conv.parse_date_only(46351), date(2026, 11, 25))

        user_warnings = [w for w in caught if issubclass(w.category, UserWarning)]
        self.assertEqual(user_warnings, [], f"Unexpected UserWarnings: {[str(w.message) for w in user_warnings]}")

    def test_t5_08_dashboard_unfiltered_export_dedup_uploader_clear_and_kpis(self) -> None:
        import views.oms_dashboard as od

        orig_csv = od.export_oms_date_only_csv
        orig_xlsx = od.export_oms_date_only_excel
        counts = {"csv": 0, "xlsx": 0}

        def _counting_csv(df: pd.DataFrame):
            counts["csv"] += 1
            return orig_csv(df)

        def _counting_xlsx(df: pd.DataFrame):
            counts["xlsx"] += 1
            return orig_xlsx(df)

        # 1. Unfiltered render serializes CSV/XLSX once; filtered render re-serializes Tab 3
        ref_df = build_reference_9_rows_df("iso")
        session = AppSession().start()
        with patch.object(od, "export_oms_date_only_csv", side_effect=_counting_csv), patch.object(
            od, "export_oms_date_only_excel", side_effect=_counting_xlsx
        ):
            session.upload_to(session.find_dashboard_uploader(), "Nov1-25.csv", to_csv_bytes(ref_df), "text/csv")
            session.assert_no_exceptions()
            self.assertEqual(counts, {"csv": 1, "xlsx": 1})

            counts["csv"] = 0
            counts["xlsx"] = 0
            session.set_date(session.find_today_date_input(), REFERENCE_TODAY)
            session.assert_no_exceptions()
            self.assertEqual(counts, {"csv": 1, "xlsx": 1})

            counts["csv"] = 0
            counts["xlsx"] = 0
            session.select_dashboard_filter("Critical")
            session.assert_no_exceptions()
            self.assertEqual(counts, {"csv": 2, "xlsx": 2})

        # 2. Clearing file_uploader resets oms_dash_file_sig to None
        up = session.find_dashboard_uploader()
        session.clear_uploader(up)
        session.assert_no_exceptions()
        self.assertIsNone(session.at.session_state["oms_dash_file_sig"])

        # 3. KPI Card 5 computes On Hold average age from On Hold rows when present
        mixed_status_df = ref_df.copy()
        # Make only EX-2026-46397 "On Hold" (Hold=0 on 2026-11-25), EX-2026-40520 "Pending Approval" (Hold=10),
        # and EX-2026-32559 "Pending CET Approval" (Hold=19)
        mixed_status_df.loc[mixed_status_df["Ticket ID"] == "EX-2026-40520", "Status"] = "Pending Approval"
        s_kpi5 = AppSession().start()
        s_kpi5.upload_to(s_kpi5.find_dashboard_uploader(), "kpi5.csv", to_csv_bytes(mixed_status_df), "text/csv")
        s_kpi5.set_date(s_kpi5.find_today_date_input(), REFERENCE_TODAY)
        s_kpi5.assert_no_exceptions()
        kpi_nov25 = s_kpi5.extract_kpi_text()
        self.assertIn("Avg On Hold age: 0.0 d", kpi_nov25)
        self.assertNotIn("9.7 d", kpi_nov25)

        s_kpi5.set_date(s_kpi5.find_today_date_input(), date(2026, 12, 5))
        s_kpi5.assert_no_exceptions()
        self.assertIn("Avg On Hold age: 10.0 d", s_kpi5.extract_kpi_text())

        # 4. KPI Card 6 falls back to Avg Ack when Resolved/Closed dates are absent
        ack_only_df = pd.DataFrame(
            [
                {
                    "Ticket ID": "EX-ACK-1",
                    "Status": "Assigned",
                    "Assigned At": "2026-11-10 09:00:00",
                    "Acknowledged On": "2026-11-12 09:00:00",
                    "Resolved At Time": "",
                    "Closed At": "",
                    "SLA Breached": "No",
                }
            ]
        )
        s_kpi6 = AppSession().start()
        s_kpi6.upload_to(s_kpi6.find_dashboard_uploader(), "ack_only.csv", to_csv_bytes(ack_only_df), "text/csv")
        s_kpi6.assert_no_exceptions()
        kpi6_text = s_kpi6.extract_kpi_text()
        self.assertIn("2.0 d", kpi6_text)
        self.assertIn("Avg Ack only (no resolution)", kpi6_text)

    def test_t5_09_compare_tab5_zone_filter_normalizes_numeric_and_padded_zones(self) -> None:
        fir_df = build_minimal_fir_df(
            [
                {
                    "OCPP ID": "CHG-Z1",
                    "Status": "Faulted",
                    "Zone": 101,
                    "Station Name": "Numeric Hub",
                },
                {
                    "OCPP ID": "CHG-Z2",
                    "Status": "Faulted",
                    "Zone": "West",
                    "Station Name": "West Hub",
                },
            ]
        )
        oms_df = pd.DataFrame(
            [
                {
                    "Ticket ID": "T-TRACKED-1",
                    "Charger": "CHG-Z1",
                    "Status": "Open",
                    "Zone": 101,
                },
                {
                    "Ticket ID": "T-UNTRACKED-INT",
                    "Charger": "CHG-EXT-101",
                    "Status": "Open",
                    "Zone": 101,
                    "Station": "Extra 101",
                },
                {
                    "Ticket ID": "T-UNTRACKED-PAD",
                    "Charger": "CHG-EXT-WEST",
                    "Status": "Open",
                    "Zone": "West ",
                    "Station": "Extra West",
                },
            ]
        )

        session = AppSession().start()
        session.switch_view(VIEW_OMS_FIR_COMPARE)
        fir_up, _ = session.find_compare_uploaders()
        session.upload_to(
            fir_up,
            "fir_zones.xlsx",
            to_xlsx_bytes(fir_df),
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        _, oms_up = session.find_compare_uploaders()
        session.upload_to(
            oms_up,
            "oms_zones.xlsx",
            to_xlsx_bytes(oms_df),
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        session.click_button("Compare Both Files")
        session.assert_no_exceptions()

        # Verify mixed-type Zone in displayed dataframes is sanitized to string dtype (no ArrowInvalid)
        for df_view in session.get_dataframes():
            if "Zone" in df_view.columns:
                self.assertIsInstance(df_view["Zone"].dtype, pd.StringDtype)

        zone_ms = session.find_multiselect("Region / City")
        session.set_multiselect(zone_ms, ["101"])
        session.assert_no_exceptions()
        self.assertIn("Extra Tickets in OMS (1)", session.get_tab_labels())
        tab5_101 = [d for d in session.get_dataframes() if "Ticket ID" in d.columns]
        self.assertEqual(len(tab5_101), 1)
        self.assertEqual(tab5_101[0]["Ticket ID"].tolist(), ["T-UNTRACKED-INT"])

        zone_ms = session.find_multiselect("Region / City")
        session.set_multiselect(zone_ms, ["West"])
        session.assert_no_exceptions()
        self.assertIn("Extra Tickets in OMS (1)", session.get_tab_labels())
        tab5_west = [d for d in session.get_dataframes() if "Ticket ID" in d.columns]
        self.assertEqual(len(tab5_west), 1)
        self.assertEqual(tab5_west[0]["Ticket ID"].tolist(), ["T-UNTRACKED-PAD"])

    def test_t5_10_reconciler_mixed_sno_without_created_at_no_crash(self) -> None:
        from core.reconciler import reconcile_incident_records

        fir_df = build_minimal_fir_df(
            [
                {"OCPP ID": "CHG-SNO-1", "Status": "Faulted", "Zone": "North"},
            ]
        )
        oms_no_ts = pd.DataFrame(
            [
                {"Sno": 10, "Ticket ID": "T-SNO-10", "Charger": "CHG-SNO-1", "Status": "Open"},
                {"Sno": "N/A", "Ticket ID": "T-SNO-NA", "Charger": "CHG-SNO-1", "Status": "Open"},
                {"Sno": 2, "Ticket ID": "T-SNO-2", "Charger": "CHG-SNO-1", "Status": "Open"},
                {"Sno": 1, "Ticket ID": "T-SNO-1", "Charger": "CHG-SNO-1", "Status": "Open"},
                {"Sno": "EXT", "Ticket ID": "T-SNO-EXT", "Charger": "CHG-EXT", "Status": "Open"},
            ]
        )

        rec = reconcile_incident_records(fir_df, oms_no_ts)
        self.assertEqual(rec.summary.active_verified, 1)
        # Numeric Sno ordering places 1 < 2 < 10 < "N/A"; "N/A" sorts after numeric values deterministically
        self.assertEqual(rec.reconciled_df.loc[0, "Latest_Ticket_ID"], "T-SNO-NA")
        self.assertNotIn("_sort_sno_num", rec.untracked_oms_df.columns)
        self.assertNotIn("_sort_sno_str", rec.untracked_oms_df.columns)

        # Also test purely numeric + string Sno where 10 > 2 > 1
        oms_num_str = pd.DataFrame(
            [
                {"Sno": 10, "Ticket ID": "T-10", "Charger": "CHG-SNO-1", "Status": "Open"},
                {"Sno": "2", "Ticket ID": "T-2", "Charger": "CHG-SNO-1", "Status": "Open"},
                {"Sno": 1, "Ticket ID": "T-1", "Charger": "CHG-SNO-1", "Status": "Open"},
            ]
        )
        rec_num = reconcile_incident_records(fir_df, oms_num_str)
        self.assertEqual(rec_num.reconciled_df.loc[0, "Latest_Ticket_ID"], "T-10")

        # UI Compare flow with mixed Sno and no Created At
        session = AppSession().start()
        session.switch_view(VIEW_OMS_FIR_COMPARE)
        fir_up, _ = session.find_compare_uploaders()
        session.upload_to(fir_up, "fir_sno.csv", to_csv_bytes(fir_df), "text/csv")
        _, oms_up = session.find_compare_uploaders()
        session.upload_to(
            oms_up,
            "oms_sno.xlsx",
            to_xlsx_bytes(oms_no_ts),
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        session.click_button("Compare Both Files")
        session.assert_no_exceptions()

    def test_t5_11_reconcile_excel_handles_mixed_object_tz_aware_datetime_and_large_float_deltas(self) -> None:
        from core.exporter import generate_reconciliation_excel
        from core.reconciler import reconcile_incident_records

        export_csv, export_xlsx = get_oms_exporters()

        fir_df = build_minimal_fir_df(
            [
                {"OCPP ID": "CHG-TZ-1", "Status": "Faulted"},
                {"OCPP ID": "CHG-TZ-2", "Status": "Faulted"},
                {"OCPP ID": "CHG-TZ-3", "Status": "Faulted"},
            ]
        )
        fir_df["Mixed_DT"] = pd.Series(
            [
                datetime(2026, 11, 25, 10, 0, tzinfo=timezone.utc),
                "pending",
                date(2026, 11, 24),
            ],
            dtype=object,
        )
        oms_df = pd.DataFrame([{"Ticket ID": "T-OTHER", "Charger": "CHG-NONE", "Status": "Open"}])

        rec = reconcile_incident_records(fir_df, oms_df)
        xlsx_bytes = generate_reconciliation_excel(rec)
        wb = load_xlsx_workbook_bytes(xlsx_bytes)
        gap_ws = wb["Ticket Generation Gaps"]
        headers = [gap_ws.cell(1, c).value for c in range(1, gap_ws.max_column + 1)]
        dt_col = headers.index("Mixed_DT") + 1
        self.assertEqual(gap_ws.cell(2, dt_col).value, "2026-11-25 10:00:00")
        self.assertEqual(gap_ws.cell(3, dt_col).value, "pending")
        self.assertEqual(gap_ws.cell(4, dt_col).value, "2026-11-24")

        # Out-of-int64-range finite float in delta column must not raise TypeError in CSV or Excel export
        huge_df = pd.DataFrame(
            [
                {"Ticket ID": "EX-HUGE-1", "Hold total time(Days)": 1e25},
                {"Ticket ID": "EX-HUGE-2", "Hold total time(Days)": 2.0},
                {"Ticket ID": "EX-HUGE-3", "Hold total time(Days)": float("inf")},
                {"Ticket ID": "EX-HUGE-4", "Hold total time(Days)": pd.NA},
            ]
        )
        csv_text = export_csv(huge_df)
        rows = list(csv.DictReader(io.StringIO(csv_text)))
        self.assertEqual(rows[0]["Hold total time(Days)"], str(int(1e25)))
        self.assertEqual(rows[1]["Hold total time(Days)"], "2")
        self.assertEqual(rows[2]["Hold total time(Days)"], "")
        self.assertEqual(rows[3]["Hold total time(Days)"], "")

        huge_xlsx = export_xlsx(huge_df)
        huge_wb = load_xlsx_workbook_bytes(huge_xlsx)
        self.assertEqual(huge_wb.active.cell(3, 2).value, 2)
        self.assertIsNone(huge_wb.active.cell(4, 2).value)

    def test_t5_12_load_oms_raw_file_preserves_single_known_header_with_keyword_data(self) -> None:
        conv = get_oms_converter()

        # 1. Multi-column Row 1 with 1 _OMS_KNOWN_HEADERS match must NOT be overridden by Row 2 data substrings
        csv_bytes = (
            b"Ticket ID,Notes,Details\n"
            b"EX-1,Charger error,Status open\n"
            b"EX-2,Charger ok,Status closed\n"
        )
        df = conv.load_oms_raw_file(csv_bytes, "single_known_header.csv")
        self.assertEqual(list(df.columns), ["Ticket ID", "Notes", "Details"])
        self.assertEqual(len(df), 2)
        self.assertEqual(df.loc[0, "Ticket ID"], "EX-1")

        # 2. Banner Row 1 above a 5-col header with 1 known OMS header + exact keywords still skips banner
        banner_csv = (
            b"Updated At,2026-11-25 10:00:00\n"
            b"Sno,Ticket ID,Charger,Station,Status\n"
            b"1,EX-BAN-3,CHG-1,St-1,Open\n"
        )
        df_banner = conv.load_oms_raw_file(banner_csv, "banner_1_known_row2.csv")
        self.assertEqual(list(df_banner.columns), ["Sno", "Ticket ID", "Charger", "Station", "Status"])
        self.assertEqual(len(df_banner), 1)

        # 3. 2-digit year 12-hour AM/PM formats (%d-%m-%y %I:%M %p and %I:%M:%S %p)
        self.assertEqual(conv.parse_date_only("25-11-26 02:30 PM"), date(2026, 11, 25))
        self.assertEqual(conv.parse_date_only("25-11-26 02:30:45 PM"), date(2026, 11, 25))
        self.assertEqual(conv.parse_date_only("25/11/26 02:30 PM"), date(2026, 11, 25))
        self.assertEqual(conv.parse_date_only("25.11.26 02:30 PM"), date(2026, 11, 25))


if __name__ == "__main__":
    unittest.main()
