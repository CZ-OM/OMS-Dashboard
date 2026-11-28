import csv
import io
import struct
import unittest
import warnings
from datetime import date, datetime, timezone, timedelta
from pathlib import Path

import numpy as np
import openpyxl
import pandas as pd

from core import (
    OMS_DATETIME_COLUMNS,
    OMS_DELTA_COLUMNS,
    OMS_DELTA_SPECS,
    RESOLVED_TIME_DATE_ONLY_COL,
    RESOLVED_TIME_RAW_COL,
    convert_oms_to_date_only,
    detect_header_row,
    export_oms_date_only_csv,
    export_oms_date_only_excel,
    load_dataset,
    load_oms_raw_file,
    parse_date_only,
    recalculate_hold_days,
    strip_to_date_str,
)
from core.loader import _parse_biff8_sst

NOV1_25_OMS_COLUMNS = [
    "Sno",                                  # 1
    "Ticket ID",                            # 2
    "Issue Category",                       # 3
    "Issue Subcategory",                    # 4
    "Station",                              # 5
    "Assigned To",                          # 6
    "Zone",                                 # 7
    "ZC-Approval",                          # 8
    "Created By",                           # 9
    "Created At",                           # 10 (dt 2)
    "Updated At",                           # 11 (dt 1)
    "Location",                             # 12
    "Priority",                             # 13
    "Status",                               # 14
    "Ticket Type",                          # 15
    "Department",                           # 16
    "Charger",                              # 17
    "Acknowledgement SLA Deadline",         # 18 (dt 4)
    "Assigned At",                          # 19 (dt 3)
    "Hold Reason",                          # 20
    "Acknowledged On",                      # 21 (dt 5)
    "Acknowledgement time taken(Days)",     # 22 (delta 1)
    "In Progress Status Time",              # 23 (dt 6)
    "SLA Breached",                         # 24
    "Hold Status Time",                     # 25 (dt 7)
    "Hold total time(Days)",                # 26 (delta 2)
    "Resolved At Time",                     # 27 (dt 8)
    RESOLVED_TIME_RAW_COL,                  # 28 (delta 3)
    "Restored At Time",                     # 29 (dt 9)
    "Restore total time(Days)",             # 30 (delta 4)
    "Closed At",                            # 31 (dt 10)
    "Total time taken (Days)",              # 32 (delta 5)
    "Resolution SLA Deadline",              # 33 (dt 11)
    "SLA Breached As per Actual (Days)",    # 34 (delta 6)
    "Approval Status",                      # 35
    "Approval Requested At Time",           # 36 (dt 12)
    "Approval Accepted At Time",            # 37 (dt 13)
    "Approval Time taken (Days)",           # 38 (delta 7)
    "CET Approval Requested At Time",       # 39 (dt 14)
    "CET Approval Accepted At Time",        # 40 (dt 15)
    "CET Approval Time taken (Days)",       # 41 (delta 8)
    "Created Month",                        # 42
]

EXPECTED_9_ROWS_DELTAS = {
    "EX-2026-49803": {
        "Acknowledgement time taken(Days)": None,
        "Hold total time(Days)": None,
        "Resolved  total time(Days)": None,
        "Restore total time(Days)": None,
        "Total time taken (Days)": None,
        "SLA Breached As per Actual (Days)": None,
        "Approval Time taken (Days)": None,
        "CET Approval Time taken (Days)": None,
    },
    "EX-2026-42245": {
        "Acknowledgement time taken(Days)": 0,
        "Hold total time(Days)": None,
        "Resolved  total time(Days)": None,
        "Restore total time(Days)": 9,
        "Total time taken (Days)": None,
        "SLA Breached As per Actual (Days)": 1,
        "Approval Time taken (Days)": None,
        "CET Approval Time taken (Days)": None,
    },
    "EX-2026-46397": {
        "Acknowledgement time taken(Days)": 2,
        "Hold total time(Days)": 0,
        "Resolved  total time(Days)": None,
        "Restore total time(Days)": None,
        "Total time taken (Days)": None,
        "SLA Breached As per Actual (Days)": 4,
        "Approval Time taken (Days)": None,
        "CET Approval Time taken (Days)": None,
    },
    "IT-2026-49625": {
        "Acknowledgement time taken(Days)": 1,
        "Hold total time(Days)": None,
        "Resolved  total time(Days)": 1,
        "Restore total time(Days)": None,
        "Total time taken (Days)": None,
        "SLA Breached As per Actual (Days)": 2,
        "Approval Time taken (Days)": None,
        "CET Approval Time taken (Days)": None,
    },
    "EX-2026-40520": {
        "Acknowledgement time taken(Days)": 0,
        "Hold total time(Days)": 10,
        "Resolved  total time(Days)": None,
        "Restore total time(Days)": 10,
        "Total time taken (Days)": None,
        "SLA Breached As per Actual (Days)": 11,
        "Approval Time taken (Days)": 1,
        "CET Approval Time taken (Days)": None,
    },
    "EX-2026-32559": {
        "Acknowledgement time taken(Days)": 0,
        "Hold total time(Days)": 19,
        "Resolved  total time(Days)": 18,
        "Restore total time(Days)": 19,
        "Total time taken (Days)": None,
        "SLA Breached As per Actual (Days)": 2,
        "Approval Time taken (Days)": -18,
        "CET Approval Time taken (Days)": 17,
    },
    "EX-2026-34312": {
        "Acknowledgement time taken(Days)": 3,
        "Hold total time(Days)": None,
        "Resolved  total time(Days)": None,
        "Restore total time(Days)": 16,
        "Total time taken (Days)": 0,
        "SLA Breached As per Actual (Days)": 6,
        "Approval Time taken (Days)": 0,
        "CET Approval Time taken (Days)": None,
    },
    "EX-2026-46325": {
        "Acknowledgement time taken(Days)": 1,
        "Hold total time(Days)": None,
        "Resolved  total time(Days)": 3,
        "Restore total time(Days)": None,
        "Total time taken (Days)": None,
        "SLA Breached As per Actual (Days)": 4,
        "Approval Time taken (Days)": 1,
        "CET Approval Time taken (Days)": 1,
    },
    "EX-2026-43053": {
        "Acknowledgement time taken(Days)": 7,
        "Hold total time(Days)": None,
        "Resolved  total time(Days)": None,
        "Restore total time(Days)": None,
        "Total time taken (Days)": None,
        "SLA Breached As per Actual (Days)": 9,
        "Approval Time taken (Days)": None,
        "CET Approval Time taken (Days)": None,
    },
}


def _build_reference_oms_df() -> pd.DataFrame:
    def _blank_row(sno: int, ticket_id: str, status: str) -> dict:
        row = {col: "" for col in NOV1_25_OMS_COLUMNS}
        row.update(
            {
                "Sno": sno,
                "Ticket ID": ticket_id,
                "Issue Category": "Hardware",
                "Issue Subcategory": "Connector Fault",
                "Station": f"Station {sno}",
                "Assigned To": f"Tech {sno}",
                "Zone": "North",
                "ZC-Approval": "No",
                "Created By": "Ops",
                "Created At": "2026-11-01 09:15:00",
                "Updated At": "2026-11-24 18:30:00",
                "Location": "Delhi",
                "Priority": "High",
                "Status": status,
                "Ticket Type": "Corrective",
                "Department": "Field Ops",
                "Charger": f"CHG-{1000 + sno}",
                "Acknowledgement SLA Deadline": "2026-11-02 09:15:00",
                "Created Month": "Nov 2026",
            }
        )
        return row

    r1 = _blank_row(1, "EX-2026-49803", "Open")
    r1.update(
        {
            "Assigned At": "",
            "Acknowledgement time taken(Days)": "colomn21-colomn19",
            "Hold total time(Days)": "today-colomn25",
            RESOLVED_TIME_RAW_COL: "colomn27-colomn19",
            "Restore total time(Days)": "colomn29-colomn19",
            "Total time taken (Days)": "colomn31-colomn19",
            "SLA Breached As per Actual (Days)": "colomn33-colomn19",
            "Approval Time taken (Days)": "colomn37-colomn36",
            "CET Approval Time taken (Days)": "colomn40-colomn39",
        }
    )

    r2 = _blank_row(2, "EX-2026-42245", "Restored")
    r2.update(
        {
            "Assigned At": "2026-11-10 08:00:00",
            "Acknowledged On": "2026-11-10 14:45:00",
            "Restored At Time": "19-11-2026 11:20:00",
            "Resolution SLA Deadline": "11/11/2026 08:00",
        }
    )

    r3 = _blank_row(3, "EX-2026-46397", "On Hold")
    r3.update(
        {
            "Assigned At": "2026-11-20 10:00:00",
            "Acknowledged On": "22 Nov 2026, 09:30:00 AM",
            "Hold Status Time": "2026-11-25 07:10:00",
            "Resolution SLA Deadline": "2026-11-24 10:00:00",
        }
    )

    r4 = _blank_row(4, "IT-2026-49625", "Resolved")
    r4.update(
        {
            "Assigned At": "2026-11-18 16:00:00",
            "Acknowledged On": "2026-11-19 09:00:00",
            "Resolved At Time": "2026-11-19 17:30:00",
            "Resolution SLA Deadline": "2026-11-20 16:00:00",
        }
    )

    r5 = _blank_row(5, "EX-2026-40520", "Pending Approval")
    r5.update(
        {
            "Assigned At": "2026-11-05 09:00:00",
            "Acknowledged On": "2026-11-05 12:00:00",
            "Hold Status Time": "2026-11-15 10:00:00",
            "Restored At Time": "2026-11-15 18:00:00",
            "Resolution SLA Deadline": "2026-11-16 09:00:00",
            "Approval Requested At Time": "2026-11-14 08:00:00",
            "Approval Accepted At Time": "2026-11-15 09:00:00",
        }
    )

    r6 = _blank_row(6, "EX-2026-32559", "Pending CET Approval")
    r6.update(
        {
            "Assigned At": "2026-11-01 09:00:00",
            "Acknowledged On": "2026-11-01 10:00:00",
            "Hold Status Time": "2026-11-06 14:00:00",
            "Resolved At Time": "2026-11-19 15:00:00",
            "Restored At Time": "2026-11-20 11:00:00",
            "Resolution SLA Deadline": "2026-11-03 09:00:00",
            "Approval Requested At Time": "2026-11-22 10:00:00",
            "Approval Accepted At Time": "2026-11-04 10:00:00",
            "CET Approval Requested At Time": "2026-11-05 09:00:00",
            "CET Approval Accepted At Time": "2026-11-22 16:00:00",
        }
    )

    r7 = _blank_row(7, "EX-2026-34312", "Closed")
    r7.update(
        {
            "Assigned At": "2026-11-02 10:00:00",
            "Acknowledged On": "2026-11-05 11:00:00",
            "Restored At Time": "2026-11-18 12:00:00",
            "Closed At": "2026-11-02 22:00:00",
            "Resolution SLA Deadline": "2026-11-08 10:00:00",
            "Approval Requested At Time": "2026-11-10 10:00:00",
            "Approval Accepted At Time": "2026-11-10 18:00:00",
        }
    )

    r8 = _blank_row(8, "EX-2026-46325", "Resolved")
    r8.update(
        {
            "Assigned At": "2026-11-12 08:00:00",
            "Acknowledged On": "2026-11-13 08:30:00",
            "Resolved At Time": "2026-11-15 11:00:00",
            "Resolution SLA Deadline": "2026-11-16 08:00:00",
            "Approval Requested At Time": "2026-11-13 09:00:00",
            "Approval Accepted At Time": "2026-11-14 09:00:00",
            "CET Approval Requested At Time": "2026-11-14 10:00:00",
            "CET Approval Accepted At Time": "2026-11-15 10:00:00",
        }
    )

    r9 = _blank_row(9, "EX-2026-43053", "Assigned")
    r9.update(
        {
            "Assigned At": "2026-11-10 09:00:00",
            "Acknowledged On": "2026-11-17 09:00:00",
            "Resolution SLA Deadline": "2026-11-19 09:00:00",
        }
    )

    return pd.DataFrame([r1, r2, r3, r4, r5, r6, r7, r8, r9], columns=NOV1_25_OMS_COLUMNS)


def _biff_record(rec_id: int, payload: bytes) -> bytes:
    return struct.pack("<HH", rec_id, len(payload)) + payload


def _biff_short_unicode(text: str) -> bytes:
    raw = text.encode("utf-16-le")
    return struct.pack("<BB", len(text), 0x01) + raw


def _biff_unicode_str(text: str) -> bytes:
    raw = text.encode("utf-16-le")
    return struct.pack("<HB", len(text), 0x01) + raw


def _biff_short_compressed(text: str, encoding: str = "latin-1") -> bytes:
    raw = text.encode(encoding)
    return struct.pack("<BB", len(raw), 0x00) + raw


def _biff_unicode_str_compressed(text: str, encoding: str = "latin-1") -> bytes:
    raw = text.encode(encoding)
    return struct.pack("<HB", len(raw), 0x00) + raw


def _wrap_ole2_v3(wb_stream: bytes) -> bytes:
    if len(wb_stream) < 4096:
        wb_stream += b"\x00" * (4096 - len(wb_stream))

    sec_size = 512
    wb_sectors = (len(wb_stream) + sec_size - 1) // sec_size
    wb_padded = wb_stream.ljust(wb_sectors * sec_size, b"\x00")

    fat_entries = [0xFFFFFFFF] * max(128, ((2 + wb_sectors + 127) // 128) * 128)
    fat_entries[0] = 0xFFFFFFFD  # FATSECT
    fat_entries[1] = 0xFFFFFFFE  # Directory end
    for i in range(wb_sectors):
        sec = 2 + i
        fat_entries[sec] = (sec + 1) if i + 1 < wb_sectors else 0xFFFFFFFE
    fat_sector = struct.pack(f"<{len(fat_entries)}I", *fat_entries[:128])

    def _dir_entry(name: str, obj_type: int, start_sec: int, size: int) -> bytes:
        name_u16 = (name + "\x00").encode("utf-16-le")
        buf = bytearray(128)
        buf[:len(name_u16)] = name_u16
        struct.pack_into("<H", buf, 0x40, len(name_u16))
        buf[0x42] = obj_type
        buf[0x43] = 1  # black
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


def _wrap_ole2_v4(wb_stream: bytes) -> bytes:
    if len(wb_stream) < 4096:
        wb_stream += b"\x00" * (4096 - len(wb_stream))

    sec_size = 4096
    wb_sectors = (len(wb_stream) + sec_size - 1) // sec_size
    wb_padded = wb_stream.ljust(wb_sectors * sec_size, b"\x00")

    entries_per_sec = sec_size // 4
    fat_entries = [0xFFFFFFFF] * entries_per_sec
    fat_entries[0] = 0xFFFFFFFD  # FATSECT
    fat_entries[1] = 0xFFFFFFFE  # Directory end
    for i in range(wb_sectors):
        sec = 2 + i
        fat_entries[sec] = (sec + 1) if i + 1 < wb_sectors else 0xFFFFFFFE
    fat_sector = struct.pack(f"<{entries_per_sec}I", *fat_entries)

    def _dir_entry(name: str, obj_type: int, start_sec: int, size: int) -> bytes:
        name_u16 = (name + "\x00").encode("utf-16-le")
        buf = bytearray(128)
        buf[:len(name_u16)] = name_u16
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
    ).ljust(sec_size, b"\x00")

    hdr = bytearray(sec_size)
    hdr[:8] = b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1"
    struct.pack_into("<HHHH", hdr, 0x18, 0x003E, 0x0004, 0xFFFE, 12)
    struct.pack_into("<H", hdr, 0x20, 6)
    struct.pack_into("<IIIIIII", hdr, 0x2C, 1, 1, 0, 4096, 0xFFFFFFFE, 0, 0xFFFFFFFE)
    struct.pack_into("<I", hdr, 0x48, 0)
    difat = [0] + [0xFFFFFFFF] * 108
    struct.pack_into("<109I", hdr, 0x4C, *difat)

    return bytes(hdr) + fat_sector + dir_sector + wb_padded


def _wrap_ole2_v3_difat_chained(wb_stream: bytes) -> bytes:
    """Build an OLE2 v3 container with 110 FAT sectors (1 chained DIFAT sector) and Workbook at sector 14000."""
    if len(wb_stream) < 4096:
        wb_stream += b"\x00" * (4096 - len(wb_stream))

    sec_size = 512
    entries_per_sec = 128
    num_fat_sectors = 110
    wb_start_sec = 14000
    wb_sectors = (len(wb_stream) + sec_size - 1) // sec_size
    wb_padded = wb_stream.ljust(wb_sectors * sec_size, b"\x00")
    total_sectors = wb_start_sec + wb_sectors

    fat_table = [0xFFFFFFFF] * (num_fat_sectors * entries_per_sec)
    for s in range(109):
        fat_table[s] = 0xFFFFFFFD  # FATSECT
    fat_table[109] = 0xFFFFFFFC  # DIFSECT
    fat_table[110] = 0xFFFFFFFD  # 110th FATSECT
    fat_table[111] = 0xFFFFFFFE  # Directory end
    for i in range(wb_sectors):
        sec = wb_start_sec + i
        fat_table[sec] = (sec + 1) if i + 1 < wb_sectors else 0xFFFFFFFE

    difat_sec_entries = [110] + [0xFFFFFFFF] * 126 + [0xFFFFFFFE]
    difat_sector = struct.pack("<128I", *difat_sec_entries)

    def _dir_entry(name: str, obj_type: int, start_sec: int, size: int) -> bytes:
        name_u16 = (name + "\x00").encode("utf-16-le")
        buf = bytearray(128)
        buf[:len(name_u16)] = name_u16
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
        + _dir_entry("Workbook", 2, wb_start_sec, len(wb_stream))
        + (b"\x00" * 256)
    )

    hdr = bytearray(sec_size)
    hdr[:8] = b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1"
    struct.pack_into("<HHHH", hdr, 0x18, 0x003E, 0x0003, 0xFFFE, 9)
    struct.pack_into("<H", hdr, 0x20, 6)
    struct.pack_into("<IIIIIII", hdr, 0x2C, num_fat_sectors, 111, 0, 4096, 0xFFFFFFFE, 0, 109)
    struct.pack_into("<I", hdr, 0x48, 1)
    struct.pack_into("<109I", hdr, 0x4C, *range(109))

    out = bytearray((1 + total_sectors) * sec_size)
    out[:sec_size] = hdr
    for i in range(109):
        chunk = struct.pack(f"<{entries_per_sec}I", *fat_table[i * entries_per_sec:(i + 1) * entries_per_sec])
        out[(1 + i) * sec_size:(2 + i) * sec_size] = chunk
    out[(1 + 109) * sec_size:(2 + 109) * sec_size] = difat_sector
    out[(1 + 110) * sec_size:(2 + 110) * sec_size] = struct.pack(
        f"<{entries_per_sec}I", *fat_table[109 * entries_per_sec:110 * entries_per_sec]
    )
    out[(1 + 111) * sec_size:(2 + 111) * sec_size] = dir_sector
    out[(1 + wb_start_sec) * sec_size:(1 + wb_start_sec) * sec_size + len(wb_padded)] = wb_padded
    return bytes(out)


def _build_synthetic_biff8_xls(rows: list[list[object]], include_filepass: bool = False) -> bytes:
    """Build a valid OLE2 + BIFF8 .xls binary containing SST, XF date format, and cell records."""
    sst_strings: list[str] = []
    sst_index: dict[str, int] = {}
    for row in rows:
        for val in row:
            if isinstance(val, str) and val not in sst_index:
                sst_index[val] = len(sst_strings)
                sst_strings.append(val)

    # XF 0 = General, XF 1 = built-in date format 14
    xf0 = struct.pack("<HHH", 0, 0, 0x0001) + b"\x00" * 14
    xf1 = struct.pack("<HHH", 0, 14, 0x0001) + b"\x00" * 14

    sst_payload = bytearray(struct.pack("<II", len(sst_strings), len(sst_strings)))
    for s in sst_strings:
        sst_payload.extend(_biff_unicode_str(s))

    # Construct globals without BOUNDSHEET first to compute exact BOUNDSHEET offset
    bof_global = _biff_record(0x0809, struct.pack("<HHHHII", 0x0600, 0x0005, 0, 0, 0, 0))
    codepage = _biff_record(0x0042, struct.pack("<H", 1200))
    datemode = _biff_record(0x0022, struct.pack("<H", 0))
    filepass = _biff_record(0x002F, b"\x00" * 6) if include_filepass else b""
    xf_recs = _biff_record(0x00E0, xf0) + _biff_record(0x00E0, xf1)
    sst_rec = _biff_record(0x00FC, bytes(sst_payload))
    eof_rec = _biff_record(0x000A, b"")

    sheet_name_bytes = _biff_short_unicode("OMSData")
    bs_len = 4 + 2 + len(sheet_name_bytes)
    sheet_offset = (
        len(bof_global)
        + len(codepage)
        + len(datemode)
        + len(filepass)
        + len(xf_recs)
        + (4 + bs_len)
        + len(sst_rec)
        + len(eof_rec)
    )
    boundsheet = _biff_record(0x0085, struct.pack("<IH", sheet_offset, 0) + sheet_name_bytes)

    ws_records = bytearray(_biff_record(0x0809, struct.pack("<HHHHII", 0x0600, 0x0010, 0, 0, 0, 0)))
    for r_idx, row in enumerate(rows):
        for c_idx, val in enumerate(row):
            if val is None:
                continue
            if isinstance(val, str):
                isst = sst_index[val]
                ws_records.extend(_biff_record(0x00FD, struct.pack("<HHHI", r_idx, c_idx, 0, isst)))
            elif isinstance(val, datetime):
                serial = (val - datetime(1899, 12, 30)).total_seconds() / 86400.0
                ws_records.extend(_biff_record(0x0203, struct.pack("<HHHd", r_idx, c_idx, 1, serial)))
            elif isinstance(val, (int, float)):
                ws_records.extend(_biff_record(0x0203, struct.pack("<HHHd", r_idx, c_idx, 0, float(val))))
    ws_records.extend(eof_rec)

    wb_stream = (
        bof_global
        + codepage
        + datemode
        + filepass
        + xf_recs
        + boundsheet
        + sst_rec
        + eof_rec
        + bytes(ws_records)
    )
    # Pad >= 4096 bytes so Workbook lives in regular FAT sectors
    if len(wb_stream) < 4096:
        wb_stream += b"\x00" * (4096 - len(wb_stream))

    sec_size = 512
    wb_sectors = (len(wb_stream) + sec_size - 1) // sec_size
    wb_padded = wb_stream.ljust(wb_sectors * sec_size, b"\x00")

    # Layout:
    # Sector 0: FAT
    # Sector 1: Directory
    # Sectors 2..(2 + wb_sectors - 1): Workbook stream
    fat_entries = [0xFFFFFFFF] * 128
    fat_entries[0] = 0xFFFFFFFD  # FATSECT
    fat_entries[1] = 0xFFFFFFFE  # Directory end
    for i in range(wb_sectors):
        sec = 2 + i
        fat_entries[sec] = (sec + 1) if i + 1 < wb_sectors else 0xFFFFFFFE
    fat_sector = struct.pack("<128I", *fat_entries)

    def _dir_entry(name: str, obj_type: int, start_sec: int, size: int) -> bytes:
        name_u16 = (name + "\x00").encode("utf-16-le")
        buf = bytearray(128)
        buf[:len(name_u16)] = name_u16
        struct.pack_into("<H", buf, 0x40, len(name_u16))
        buf[0x42] = obj_type
        buf[0x43] = 1  # black
        struct.pack_into("<III", buf, 0x44, 0xFFFFFFFF, 0xFFFFFFFF, 0xFFFFFFFF)
        if obj_type == 5:
            struct.pack_into("<I", buf, 0x4C, 1)  # child = Workbook
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


class TestOMSConverter(unittest.TestCase):
    def test_schema_constants(self):
        self.assertEqual(len(OMS_DATETIME_COLUMNS), 15)
        self.assertEqual(RESOLVED_TIME_RAW_COL, "Resolved  total time(Days:Hours)")
        self.assertEqual(RESOLVED_TIME_DATE_ONLY_COL, "Resolved  total time(Days)")
        self.assertEqual(len(OMS_DELTA_SPECS), 8)
        self.assertEqual(len(OMS_DELTA_COLUMNS), 8)
        self.assertEqual(OMS_DELTA_SPECS[1], ("Hold total time(Days)", "Hold Status Time", None))

    def test_convert_9_reference_rows_exact_deltas_and_schema(self):
        raw_df = _build_reference_oms_df()
        raw_snapshot = raw_df.copy(deep=True)

        converted = convert_oms_to_date_only(raw_df, today_date=date(2026, 11, 25))

        # Caller DataFrame must not be mutated
        pd.testing.assert_frame_equal(raw_df, raw_snapshot)

        # 42-column order preserved with exact rename of Resolved total time
        self.assertEqual(len(converted.columns), 42)
        expected_cols = [
            RESOLVED_TIME_DATE_ONLY_COL if c == RESOLVED_TIME_RAW_COL else c
            for c in NOV1_25_OMS_COLUMNS
        ]
        self.assertEqual(list(converted.columns), expected_cols)
        self.assertNotIn(RESOLVED_TIME_RAW_COL, converted.columns)

        # All 15 datetime columns stripped to YYYY-MM-DD or ""
        for dt_col in OMS_DATETIME_COLUMNS:
            for val in converted[dt_col]:
                self.assertIsInstance(val, str)
                if val != "":
                    self.assertRegex(val, r"^\d{4}-\d{2}-\d{2}$")

        # All 8 delta columns use nullable Int64 and match exact reference outputs
        for delta_col in OMS_DELTA_COLUMNS:
            self.assertEqual(str(converted[delta_col].dtype), "Int64")

        for _, row in converted.iterrows():
            tid = row["Ticket ID"]
            expected = EXPECTED_9_ROWS_DELTAS[tid]
            for delta_col, exp_val in expected.items():
                actual = row[delta_col]
                if exp_val is None:
                    self.assertTrue(pd.isna(actual), f"{tid} {delta_col}: expected NA, got {actual!r}")
                else:
                    self.assertFalse(pd.isna(actual), f"{tid} {delta_col}: expected {exp_val}, got NA")
                    self.assertEqual(int(actual), exp_val, f"{tid} {delta_col} mismatch")

    def test_dynamic_today_date_and_recalculate_hold_days(self):
        raw_df = _build_reference_oms_df()

        default_conv = convert_oms_to_date_only(raw_df)
        self.assertEqual(default_conv.attrs.get("today_date"), date.today().strftime("%Y-%m-%d"))

        conv_nov25 = convert_oms_to_date_only(raw_df, today_date=date(2026, 11, 25))
        recalc = recalculate_hold_days(conv_nov25, today_date=date(2026, 11, 30))

        # Non-mutating check
        self.assertEqual(
            int(conv_nov25.loc[conv_nov25["Ticket ID"] == "EX-2026-46397", "Hold total time(Days)"].iloc[0]),
            0,
        )

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
        for _, row in recalc.iterrows():
            tid = row["Ticket ID"]
            exp = expected_hold[tid]
            val = row["Hold total time(Days)"]
            if exp is None:
                self.assertTrue(pd.isna(val))
            else:
                self.assertEqual(int(val), exp)
        self.assertEqual(recalc.attrs.get("today_date"), "2026-11-30")

    def test_date_parsing_edge_cases_and_zero_warnings(self):
        with warnings.catch_warnings(record=True) as caught:
            warnings.simplefilter("always")

            # ISO vs DD-MM-YYYY must NEVER swap month and day
            self.assertEqual(parse_date_only("2026-11-01"), date(2026, 11, 1))
            self.assertEqual(strip_to_date_str("2026-11-01"), "2026-11-01")
            self.assertEqual(parse_date_only("01-11-2026"), date(2026, 11, 1))
            self.assertEqual(strip_to_date_str("01-11-2026"), "2026-11-01")
            self.assertEqual(strip_to_date_str("01/11/2026 08:45:00"), "2026-11-01")
            self.assertEqual(strip_to_date_str("11 Nov 2026, 12:11:57 PM"), "2026-11-11")

            # Timezone-aware ISO and Timestamp
            self.assertEqual(strip_to_date_str("2026-11-25T10:30:00Z"), "2026-11-25")
            self.assertEqual(strip_to_date_str("2026-11-25 14:22:10+05:30"), "2026-11-25")
            tz_ts = pd.Timestamp("2026-11-25 23:30:00", tz=timezone(timedelta(hours=5, minutes=30)))
            self.assertEqual(strip_to_date_str(tz_ts), "2026-11-25")
            self.assertEqual(strip_to_date_str(datetime(2026, 11, 25, 16, 40)), "2026-11-25")
            self.assertEqual(strip_to_date_str(date(2026, 11, 25)), "2026-11-25")

            # Excel serial numbers
            self.assertEqual(strip_to_date_str(46351), "2026-11-25")
            self.assertEqual(strip_to_date_str("46351.5"), "2026-11-25")
            self.assertEqual(strip_to_date_str(100), "")
            self.assertEqual(strip_to_date_str(90000), "")

            # Booleans, nulls, sentinels, formula-hint strings
            for bad in (None, True, False, pd.NA, pd.NaT, float("nan"), "", "-", "--", "N/A", "NA", "NULL", "NaT", "colomn21-colomn19", "today-colomn25", "=A1-B1"):
                self.assertIsNone(parse_date_only(bad), f"Expected None for {bad!r}")
                self.assertEqual(strip_to_date_str(bad), "", f"Expected '' for {bad!r}")

        user_warnings = [w for w in caught if issubclass(w.category, UserWarning)]
        self.assertEqual(user_warnings, [], f"Unexpected UserWarnings: {[str(w.message) for w in user_warnings]}")

    def test_subset_dump_oms_dump_xlsx(self):
        dump_path = Path(r"D:\Comperator\DUMP\OMS DUMP.xlsx")
        if not dump_path.exists():
            self.skipTest("Sample OMS DUMP.xlsx not present on disk")

        raw_df = load_oms_raw_file(dump_path, filename="OMS DUMP.xlsx")
        self.assertEqual(len(raw_df.columns), 24)
        converted = convert_oms_to_date_only(raw_df, today_date=date(2026, 11, 25))
        self.assertEqual(len(converted), len(raw_df))
        self.assertEqual(list(converted.columns), list(raw_df.columns))
        for col in ("Created At", "Updated At", "Acknowledgement SLA Deadline", "Resolution SLA Deadline", "Closed At"):
            non_empty = [v for v in converted[col] if v != ""]
            self.assertTrue(len(non_empty) > 0)
            for v in non_empty[:20]:
                self.assertRegex(v, r"^\d{4}-\d{2}-\d{2}$")

    def test_multi_format_csv_xlsx_and_xls_ingestion(self):
        raw_df = _build_reference_oms_df()

        # 1. CSV (UTF-8 and cp1252)
        csv_bytes = raw_df.to_csv(index=False).encode("utf-8")
        loaded_csv = load_oms_raw_file(io.BytesIO(csv_bytes), filename="nov1.csv")
        self.assertEqual(list(loaded_csv.columns), NOV1_25_OMS_COLUMNS)
        conv_csv = convert_oms_to_date_only(loaded_csv, today_date=date(2026, 11, 25))
        self.assertEqual(int(conv_csv.loc[5, "Approval Time taken (Days)"]), -18)

        cp1252_csv = "Ticket ID,Assigned At,Acknowledged On\nT1\u2013A,2026-11-10,2026-11-12\n".encode("cp1252")
        loaded_cp = load_oms_raw_file(cp1252_csv, filename="cp.csv")
        self.assertEqual(loaded_cp.loc[0, "Ticket ID"], "T1\u2013A")

        # 2. XLSX
        xlsx_buf = io.BytesIO()
        raw_df.to_excel(xlsx_buf, index=False)
        xlsx_buf.seek(0)
        loaded_xlsx = load_oms_raw_file(xlsx_buf, filename="nov1.xlsx")
        self.assertIn(RESOLVED_TIME_RAW_COL, loaded_xlsx.columns)

        # 3. Misnamed OOXML as .xls
        xlsx_buf.seek(0)
        loaded_misnamed = load_oms_raw_file(xlsx_buf.getvalue(), filename="misnamed.xls")
        self.assertEqual(list(loaded_misnamed.columns), NOV1_25_OMS_COLUMNS)

        # 4. Synthetic BIFF8 binary .xls (row 1 header AND banner row + report_date)
        biff_rows = [
            ["Ticket ID", "Assigned At", "Acknowledged On", RESOLVED_TIME_RAW_COL, "Hold Status Time"],
            ["EX-1", datetime(2026, 11, 10, 9, 30), datetime(2026, 11, 12, 14, 0), "2:04", datetime(2026, 11, 20)],
        ]
        biff_bytes = _build_synthetic_biff8_xls(biff_rows)
        loaded_biff = load_oms_raw_file(biff_bytes, filename="oms.xls")
        self.assertIn(RESOLVED_TIME_RAW_COL, loaded_biff.columns)
        conv_biff = convert_oms_to_date_only(loaded_biff, today_date=date(2026, 11, 25))
        self.assertEqual(int(conv_biff.loc[0, "Acknowledgement time taken(Days)"]), 2)
        self.assertEqual(int(conv_biff.loc[0, "Hold total time(Days)"]), 5)

        biff_banner_rows = [
            ["Report Generated On: 25-11-2026", None, None, None],
            ["Sno", "Ticket ID", "Charger", "Status"],
            [1, "EX-99", "CHG-1", "Open"],
        ]
        biff_banner_bytes = _build_synthetic_biff8_xls(biff_banner_rows)
        self.assertEqual(detect_header_row(biff_banner_bytes, filename="banner.xls"), 2)
        loaded_banner = load_dataset(biff_banner_bytes, filename="banner.xls")
        self.assertEqual(loaded_banner.attrs.get("report_date"), "25-11-2026")
        self.assertEqual(loaded_banner.loc[0, "Ticket ID"], "EX-99")

        # Encrypted BIFF8 (FILEPASS) raises clear ValueError
        encrypted_biff = _build_synthetic_biff8_xls(biff_rows, include_filepass=True)
        with self.assertRaises(ValueError):
            load_dataset(encrypted_biff, filename="encrypted.xls")

        # 5. SpreadsheetML 2003 XML .xls
        xml_xls = f"""<?xml version="1.0"?>
        <Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
                  xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">
          <Worksheet ss:Name="Sheet1">
            <Table>
              <Row>
                <Cell><Data ss:Type="String">Ticket ID</Data></Cell>
                <Cell><Data ss:Type="String">Assigned At</Data></Cell>
                <Cell ss:Index="4"><Data ss:Type="String">{RESOLVED_TIME_RAW_COL}</Data></Cell>
                <Cell><Data ss:Type="String">Resolved At Time</Data></Cell>
              </Row>
              <Row>
                <Cell><Data ss:Type="String">EX-XML-1</Data></Cell>
                <Cell><Data ss:Type="String">2026-11-10 08:00:00</Data></Cell>
                <Cell ss:Index="4"><Data ss:Type="String">ignored</Data></Cell>
                <Cell><Data ss:Type="String">2026-11-14 09:00:00</Data></Cell>
              </Row>
            </Table>
          </Worksheet>
        </Workbook>""".encode("utf-8")
        loaded_xml = load_oms_raw_file(xml_xls, filename="report.xls")
        conv_xml = convert_oms_to_date_only(loaded_xml, today_date=date(2026, 11, 25))
        self.assertEqual(int(conv_xml.loc[0, RESOLVED_TIME_DATE_ONLY_COL]), 4)

        # 6. HTML table .xls preserving double spaces via &nbsp; / literal spaces
        html_xls = (
            "<html><body><table>"
            "<tr><th>Ticket ID</th><th>Assigned At</th><th>Acknowledged On</th>"
            "<th>Resolved&nbsp; total time(Days:Hours)</th></tr>"
            "<tr><td>EX-HTML-1</td><td>2026-11-05</td><td>2026-11-08</td><td>3:00</td></tr>"
            "</table></body></html>"
        ).encode("utf-8")
        loaded_html = load_oms_raw_file(html_xls, filename="export.xls")
        self.assertIn(RESOLVED_TIME_RAW_COL, loaded_html.columns)
        conv_html = convert_oms_to_date_only(loaded_html, today_date=date(2026, 11, 25))
        self.assertEqual(int(conv_html.loc[0, "Acknowledgement time taken(Days)"]), 3)

        # 7. UTF-16LE TSV disguised as .xls
        tsv_text = f"Ticket ID\tAssigned At\tAcknowledged On\t{RESOLVED_TIME_RAW_COL}\nEX-TSV-1\t2026-11-01\t2026-11-03\t2:00\n"
        tsv_bytes = b"\xff\xfe" + tsv_text.encode("utf-16-le")
        loaded_tsv = load_oms_raw_file(tsv_bytes, filename="tsv.xls")
        conv_tsv = convert_oms_to_date_only(loaded_tsv, today_date=date(2026, 11, 25))
        self.assertEqual(int(conv_tsv.loc[0, "Acknowledgement time taken(Days)"]), 2)

        # 8. Empty / corrupt .xls raises ValueError in load_dataset while detect_header_row stays safe
        self.assertEqual(detect_header_row(b"", filename="empty.xls"), 1)
        with self.assertRaises(ValueError):
            load_dataset(b"", filename="empty.xls")
        with self.assertRaises(ValueError):
            load_dataset(b"\x01\x02\x03\x04\x05corrupt-binary", filename="bad.xls")

    def test_csv_and_excel_export_round_trip(self):
        raw_df = _build_reference_oms_df()
        converted = convert_oms_to_date_only(raw_df, today_date=date(2026, 11, 25))

        # CSV round-trip
        csv_text = export_oms_date_only_csv(converted)
        rows = list(csv.DictReader(io.StringIO(csv_text)))
        self.assertEqual(len(rows), 9)
        self.assertEqual(list(rows[0].keys()), list(converted.columns))
        self.assertEqual(rows[0]["Acknowledgement time taken(Days)"], "")
        self.assertEqual(rows[1]["Acknowledgement time taken(Days)"], "0")
        self.assertEqual(rows[5]["Approval Time taken (Days)"], "-18")
        self.assertEqual(rows[5]["Hold total time(Days)"], "19")
        self.assertNotIn(".0", rows[5]["Approval Time taken (Days)"])

        # Excel round-trip
        xlsx_bytes = export_oms_date_only_excel(converted)
        wb = openpyxl.load_workbook(io.BytesIO(xlsx_bytes))
        ws = wb.active
        sheet_rows = list(ws.iter_rows(values_only=True))
        self.assertEqual(len(sheet_rows), 10)  # 1 header + 9 data rows
        headers = list(sheet_rows[0])
        self.assertEqual(headers, list(converted.columns))

        ack_idx = headers.index("Acknowledgement time taken(Days)")
        hold_idx = headers.index("Hold total time(Days)")
        app_idx = headers.index("Approval Time taken (Days)")
        assigned_idx = headers.index("Assigned At")

        # Row 1 (EX-2026-49803): blank date and blank deltas must be None
        self.assertIsNone(sheet_rows[1][assigned_idx])
        self.assertIsNone(sheet_rows[1][ack_idx])
        self.assertIsNone(sheet_rows[1][hold_idx])

        # Row 2 (EX-2026-42245): exact Python int 0
        self.assertIs(type(sheet_rows[2][ack_idx]), int)
        self.assertEqual(sheet_rows[2][ack_idx], 0)
        self.assertEqual(sheet_rows[2][assigned_idx], "2026-11-10")

        # Row 6 (EX-2026-32559): exact Python int -18 and 19
        self.assertIs(type(sheet_rows[6][app_idx]), int)
        self.assertEqual(sheet_rows[6][app_idx], -18)
        self.assertIs(type(sheet_rows[6][hold_idx]), int)
        self.assertEqual(sheet_rows[6][hold_idx], 19)

        # pandas read_excel also reads it cleanly
        df_readback = pd.read_excel(io.BytesIO(xlsx_bytes))
        self.assertEqual(len(df_readback), 9)

    def test_biff8_codepage_1200_compressed_strings(self):
        def _build_compressed_biff8(cp_num: int, encoding: str) -> bytes:
            bof_global = _biff_record(0x0809, struct.pack("<HHHHII", 0x0600, 0x0005, 0, 0, 0, 0))
            codepage_rec = _biff_record(0x0042, struct.pack("<H", cp_num))
            datemode_rec = _biff_record(0x0022, struct.pack("<H", 0))
            fmt_rec = _biff_record(
                0x041E,
                struct.pack("<H", 164) + _biff_unicode_str_compressed("yyyy-mm-dd", encoding),
            )
            xf0 = struct.pack("<HHH", 0, 0, 0x0001) + b"\x00" * 14
            xf1 = struct.pack("<HHH", 0, 164, 0x0001) + b"\x00" * 14
            xf_recs = _biff_record(0x00E0, xf0) + _biff_record(0x00E0, xf1)

            sst_strings = ["Ticket ID", "Assigned At", "City", "FormulaCol"]
            sst_payload = bytearray(struct.pack("<II", len(sst_strings), len(sst_strings)))
            for s in sst_strings:
                sst_payload.extend(_biff_unicode_str_compressed(s, encoding))
            sst_rec = _biff_record(0x00FC, bytes(sst_payload))
            eof_rec = _biff_record(0x000A, b"")

            sheet_name_bytes = _biff_short_compressed("SheetOne", encoding)
            bs_len = 4 + 2 + len(sheet_name_bytes)
            sheet_offset = (
                len(bof_global)
                + len(codepage_rec)
                + len(datemode_rec)
                + len(fmt_rec)
                + len(xf_recs)
                + (4 + bs_len)
                + len(sst_rec)
                + len(eof_rec)
            )
            boundsheet = _biff_record(0x0085, struct.pack("<IH", sheet_offset, 0) + sheet_name_bytes)

            ws = bytearray(_biff_record(0x0809, struct.pack("<HHHHII", 0x0600, 0x0010, 0, 0, 0, 0)))
            for c_idx in range(4):
                ws.extend(_biff_record(0x00FD, struct.pack("<HHHI", 0, c_idx, 0, c_idx)))
            # Row 1: LABEL (0x0204), NUMBER date (XF=1), RSTRING (0x00D6), FORMULA+STRING (0x0207)
            ws.extend(
                _biff_record(
                    0x0204,
                    struct.pack("<HHH", 1, 0, 0) + _biff_unicode_str_compressed("EX-99", encoding),
                )
            )
            serial = (datetime(2026, 11, 20) - datetime(1899, 12, 30)).days
            ws.extend(_biff_record(0x0203, struct.pack("<HHHd", 1, 1, 1, float(serial))))
            city_val = "Café" if encoding == "latin-1" else "City\u2013A"
            ws.extend(
                _biff_record(
                    0x00D6,
                    struct.pack("<HHH", 1, 2, 0) + _biff_unicode_str_compressed(city_val, encoding),
                )
            )
            # FORMULA kind=0 (string result follows in 0x0207)
            res_bytes = struct.pack("<B5sH", 0, b"\x00" * 5, 0xFFFF)
            ws.extend(_biff_record(0x0006, struct.pack("<HHH", 1, 3, 0) + res_bytes + b"\x00" * 8))
            ws.extend(_biff_record(0x0207, _biff_unicode_str_compressed("CachedStr", encoding)))
            ws.extend(eof_rec)

            wb_stream = (
                bof_global
                + codepage_rec
                + datemode_rec
                + fmt_rec
                + xf_recs
                + boundsheet
                + sst_rec
                + eof_rec
                + bytes(ws)
            )
            return _wrap_ole2_v3(wb_stream)

        xls_1200 = _build_compressed_biff8(1200, "latin-1")
        df_1200 = load_dataset(xls_1200, filename="cp1200.xls", sheet_name="SheetOne", header_row=1)
        self.assertEqual(list(df_1200.columns), ["Ticket ID", "Assigned At", "City", "FormulaCol"])
        self.assertEqual(df_1200.loc[0, "Ticket ID"], "EX-99")
        self.assertEqual(df_1200.loc[0, "Assigned At"], pd.Timestamp("2026-11-20"))
        self.assertEqual(df_1200.loc[0, "City"], "Café")
        self.assertEqual(df_1200.loc[0, "FormulaCol"], "CachedStr")

        # Regression: CODEPAGE=1252 preserves Windows-1252 en-dash (\u2013)
        xls_1252 = _build_compressed_biff8(1252, "cp1252")
        df_1252 = load_dataset(xls_1252, filename="cp1252.xls", sheet_name="SheetOne", header_row=1)
        self.assertEqual(df_1252.loc[0, "City"], "City\u2013A")

    def test_unpadded_banner_csv_and_delimited_xls(self):
        # 1. Single unpadded 1-column banner row in CSV
        csv_banner = b"OMS Report 25-11-2026\nTicket ID,Assigned At,Status\nEX-4,2026-11-20,Open\n"
        self.assertEqual(detect_header_row(csv_banner, filename="b.csv"), 2)
        df_banner = load_oms_raw_file(csv_banner, filename="b.csv")
        self.assertEqual(df_banner.shape, (1, 3))
        self.assertEqual(list(df_banner.columns), ["Ticket ID", "Assigned At", "Status"])
        self.assertEqual(df_banner.attrs.get("report_date"), "25-11-2026")
        self.assertEqual(df_banner.loc[0, "Ticket ID"], "EX-4")

        # 2. Multi-row ragged banner with blank line before header (skiprows alignment)
        ragged_csv = (
            b"Company Banner\n"
            b"\n"
            b"Date,25-11-2026\n"
            b"Ticket ID,Assigned At,Status\n"
            b"EX-5,2026-11-21,Closed\n"
        )
        self.assertEqual(detect_header_row(ragged_csv, filename="ragged.csv"), 4)
        df_ragged = load_dataset(ragged_csv, filename="ragged.csv")
        self.assertEqual(df_ragged.shape, (1, 3))
        self.assertEqual(list(df_ragged.columns), ["Ticket ID", "Assigned At", "Status"])
        self.assertEqual(df_ragged.attrs.get("report_date"), "25-11-2026")
        self.assertEqual(df_ragged.loc[0, "Ticket ID"], "EX-5")

        # 3. UTF-8 BOM + banner CSV must strip BOM cleanly
        bom_csv = b"\xef\xbb\xbfTicket ID,Assigned At,Status\nEX-6,2026-11-22,Open\n"
        df_bom = load_oms_raw_file(bom_csv, filename="bom.csv")
        self.assertEqual(list(df_bom.columns), ["Ticket ID", "Assigned At", "Status"])
        self.assertNotIn("\ufeff", df_bom.columns[0])

        bom_banner_csv = b"\xef\xbb\xbfOMS Banner 25-11-2026\nTicket ID,Assigned At,Status\nEX-7,2026-11-22,Open\n"
        df_bom_banner = load_oms_raw_file(bom_banner_csv, filename="bom_banner.csv")
        self.assertEqual(list(df_bom_banner.columns), ["Ticket ID", "Assigned At", "Status"])
        self.assertEqual(df_bom_banner.attrs.get("report_date"), "25-11-2026")

        # 4. Delimited .xls (ASCII TSV and UTF-16LE BOM TSV) with unpadded 1-column banner
        tsv_banner = (
            f"OMS Report 25-11-2026\n"
            f"Ticket ID\tAssigned At\t{RESOLVED_TIME_RAW_COL}\n"
            f"EX-TSV-2\t2026-11-20\t1:00\n"
        )
        ascii_tsv_bytes = tsv_banner.encode("utf-8")
        self.assertEqual(detect_header_row(ascii_tsv_bytes, filename="tsv_banner.xls"), 2)
        df_tsv = load_oms_raw_file(ascii_tsv_bytes, filename="tsv_banner.xls")
        self.assertEqual(list(df_tsv.columns), ["Ticket ID", "Assigned At", RESOLVED_TIME_RAW_COL])
        self.assertEqual(df_tsv.attrs.get("report_date"), "25-11-2026")

        u16_tsv_bytes = b"\xff\xfe" + tsv_banner.encode("utf-16-le")
        self.assertEqual(detect_header_row(u16_tsv_bytes, filename="u16_banner.xls"), 2)
        df_u16 = load_oms_raw_file(u16_tsv_bytes, filename="u16_banner.xls")
        self.assertEqual(list(df_u16.columns), ["Ticket ID", "Assigned At", RESOLVED_TIME_RAW_COL])

    def test_biff8_formula_with_shrfmla_array_table_and_empty_continue(self):
        # 1. Direct _parse_biff8_sst with an empty CONTINUE chunk
        s1 = _biff_unicode_str_compressed("Alpha")
        s2_part1 = struct.pack("<HB", 5, 0x00) + b"Be"
        s2_part2 = b"\x00" + b"ta!"  # CONTINUE cont_flag=0x00 + "ta!"
        chunk0 = struct.pack("<II", 2, 2) + s1 + s2_part1
        parsed = _parse_biff8_sst([chunk0, b"", s2_part2], codepage="latin-1")
        self.assertEqual(parsed, ["Alpha", "Beta!"])

        # 2. Full BIFF8 workbook with SHRFMLA (0x04BC), ARRAY (0x0221), TABLE (0x0236), and empty CONTINUE
        bof_global = _biff_record(0x0809, struct.pack("<HHHHII", 0x0600, 0x0005, 0, 0, 0, 0))
        codepage_rec = _biff_record(0x0042, struct.pack("<H", 1200))
        xf0 = struct.pack("<HHH", 0, 0, 0x0001) + b"\x00" * 14
        xf_rec = _biff_record(0x00E0, xf0)

        hdrs = ["Ticket ID", "ShrCol", "ArrCol", "TblCol"]
        sst_p0 = bytearray(struct.pack("<II", len(hdrs), len(hdrs)))
        for h in hdrs[:2]:
            sst_p0.extend(_biff_unicode_str_compressed(h))
        # Split 3rd string across CONTINUE with an empty CONTINUE in between
        sst_p0.extend(struct.pack("<HB", 6, 0x00) + b"Arr")
        sst_cont_empty = _biff_record(0x003C, b"")
        sst_cont_data = _biff_record(0x003C, b"\x00" + b"Col" + _biff_unicode_str_compressed("TblCol"))
        sst_recs = _biff_record(0x00FC, bytes(sst_p0)) + sst_cont_empty + sst_cont_data
        eof_rec = _biff_record(0x000A, b"")

        sname = _biff_short_compressed("S1")
        sheet_offset = (
            len(bof_global)
            + len(codepage_rec)
            + len(xf_rec)
            + (4 + 6 + len(sname))
            + len(sst_recs)
            + len(eof_rec)
        )
        boundsheet = _biff_record(0x0085, struct.pack("<IH", sheet_offset, 0) + sname)

        ws = bytearray(_biff_record(0x0809, struct.pack("<HHHHII", 0x0600, 0x0010, 0, 0, 0, 0)))
        for c_idx in range(4):
            ws.extend(_biff_record(0x00FD, struct.pack("<HHHI", 0, c_idx, 0, c_idx)))
        ws.extend(_biff_record(0x0204, struct.pack("<HHH", 1, 0, 0) + _biff_unicode_str_compressed("EX-F1")))

        str_res = struct.pack("<B5sH", 0, b"\x00" * 5, 0xFFFF)
        for c_idx, mid_id, val in (
            (1, 0x04BC, "SharedVal"),
            (2, 0x0221, "ArrayVal"),
            (3, 0x0236, "TableVal"),
        ):
            ws.extend(_biff_record(0x0006, struct.pack("<HHH", 1, c_idx, 0) + str_res + b"\x00" * 8))
            ws.extend(_biff_record(mid_id, b"\x00" * 8))
            ws.extend(_biff_record(0x0207, _biff_unicode_str_compressed(val)))
        ws.extend(eof_rec)

        wb_stream = (
            bof_global + codepage_rec + xf_rec + boundsheet + sst_recs + eof_rec + bytes(ws)
        )
        xls_bytes = _wrap_ole2_v3(wb_stream)
        df = load_oms_raw_file(xls_bytes, filename="formulas.xls")
        self.assertEqual(list(df.columns), hdrs)
        self.assertEqual(df.loc[0, "ShrCol"], "SharedVal")
        self.assertEqual(df.loc[0, "ArrCol"], "ArrayVal")
        self.assertEqual(df.loc[0, "TblCol"], "TableVal")

    def test_ole2_v4_4096_sectors_and_difat_chaining(self):
        bof_global = _biff_record(0x0809, struct.pack("<HHHHII", 0x0600, 0x0005, 0, 0, 0, 0))
        codepage_rec = _biff_record(0x0042, struct.pack("<H", 1200))
        xf_rec = _biff_record(0x00E0, struct.pack("<HHH", 0, 0, 0x0001) + b"\x00" * 14)
        hdrs = ["Ticket ID", "Assigned At", "Status"]
        sst_payload = bytearray(struct.pack("<II", len(hdrs), len(hdrs)))
        for h in hdrs:
            sst_payload.extend(_biff_unicode_str_compressed(h))
        sst_rec = _biff_record(0x00FC, bytes(sst_payload))
        eof_rec = _biff_record(0x000A, b"")

        sname = _biff_short_compressed("Data")
        sheet_offset = (
            len(bof_global) + len(codepage_rec) + len(xf_rec) + (4 + 6 + len(sname)) + len(sst_rec) + len(eof_rec)
        )
        boundsheet = _biff_record(0x0085, struct.pack("<IH", sheet_offset, 0) + sname)

        ws = bytearray(_biff_record(0x0809, struct.pack("<HHHHII", 0x0600, 0x0010, 0, 0, 0, 0)))
        for c_idx in range(3):
            ws.extend(_biff_record(0x00FD, struct.pack("<HHHI", 0, c_idx, 0, c_idx)))
        ws.extend(_biff_record(0x0204, struct.pack("<HHH", 1, 0, 0) + _biff_unicode_str_compressed("EX-OLE")))
        ws.extend(_biff_record(0x0204, struct.pack("<HHH", 1, 1, 0) + _biff_unicode_str_compressed("2026-11-20")))
        ws.extend(_biff_record(0x0204, struct.pack("<HHH", 1, 2, 0) + _biff_unicode_str_compressed("Open")))
        ws.extend(eof_rec)

        wb_stream = bof_global + codepage_rec + xf_rec + boundsheet + sst_rec + eof_rec + bytes(ws)

        # 1. OLE2 v4 (4096-byte sectors)
        v4_bytes = _wrap_ole2_v4(wb_stream)
        df_v4 = load_oms_raw_file(v4_bytes, filename="v4.xls")
        self.assertEqual(list(df_v4.columns), hdrs)
        self.assertEqual(df_v4.loc[0, "Ticket ID"], "EX-OLE")

        # 2. OLE2 v3 with 110 FAT sectors + 1 chained DIFAT sector
        difat_bytes = _wrap_ole2_v3_difat_chained(wb_stream)
        df_difat = load_oms_raw_file(difat_bytes, filename="difat.xls")
        self.assertEqual(list(df_difat.columns), hdrs)
        self.assertEqual(df_difat.loc[0, "Ticket ID"], "EX-OLE")

    def test_non_seekable_streams(self):
        class NonSeekableRawStream(io.RawIOBase):
            def __init__(self, data: bytes) -> None:
                self._buf = io.BytesIO(data)

            def readable(self) -> bool:
                return True

            def seekable(self) -> bool:
                return False

            def read(self, size: int = -1) -> bytes:
                return self._buf.read(size)

        class ReadOnlyStream:
            def __init__(self, data: bytes) -> None:
                self._buf = io.BytesIO(data)

            def read(self, size: int = -1) -> bytes:
                return self._buf.read(size)

        csv_banner = b"OMS Report 25-11-2026\nTicket ID,Assigned At,Status\nEX-NS1,2026-11-20,Open\n"
        self.assertEqual(detect_header_row(NonSeekableRawStream(csv_banner), filename="ns.csv"), 2)
        df_csv_ns = load_oms_raw_file(NonSeekableRawStream(csv_banner), filename="ns.csv")
        self.assertEqual(list(df_csv_ns.columns), ["Ticket ID", "Assigned At", "Status"])
        self.assertEqual(df_csv_ns.loc[0, "Ticket ID"], "EX-NS1")

        df_csv_ro = load_dataset(ReadOnlyStream(csv_banner), filename="ro.csv")
        self.assertEqual(df_csv_ro.attrs.get("report_date"), "25-11-2026")
        self.assertEqual(df_csv_ro.loc[0, "Ticket ID"], "EX-NS1")

        # XLSX on non-seekable stream
        xlsx_buf = io.BytesIO()
        pd.DataFrame([{"Ticket ID": "EX-XLSX-NS", "Assigned At": "2026-11-20", "Status": "Open"}]).to_excel(
            xlsx_buf, index=False
        )
        xlsx_bytes = xlsx_buf.getvalue()
        df_xlsx_ns = load_oms_raw_file(NonSeekableRawStream(xlsx_bytes), filename="ns.xlsx")
        self.assertEqual(df_xlsx_ns.loc[0, "Ticket ID"], "EX-XLSX-NS")
        df_xlsx_ro = load_dataset(ReadOnlyStream(xlsx_bytes), filename="ro.xlsx")
        self.assertEqual(df_xlsx_ro.loc[0, "Ticket ID"], "EX-XLSX-NS")

        # XLS on non-seekable stream
        biff_bytes = _build_synthetic_biff8_xls(
            [["Ticket ID", "Assigned At", "Status"], ["EX-XLS-NS", "2026-11-20", "Open"]]
        )
        df_xls_ns = load_oms_raw_file(NonSeekableRawStream(biff_bytes), filename="ns.xls")
        self.assertEqual(df_xls_ns.loc[0, "Ticket ID"], "EX-XLS-NS")
        df_xls_ro = load_dataset(ReadOnlyStream(biff_bytes), filename="ro.xls")
        self.assertEqual(df_xls_ro.loc[0, "Ticket ID"], "EX-XLS-NS")

    def test_parse_date_only_timezone_abbreviations_and_short_token_rejection(self):
        with warnings.catch_warnings(record=True) as caught:
            warnings.simplefilter("always")

            tz_valid = [
                "2026-11-25 14:30:00 IST",
                "2026-11-25 14:30:00 UTC",
                "2026-11-25 14:30:00 GMT",
                "2026-11-25 14:30:00 EST",
                "2026-11-25 14:30:00 PST",
                "2026-11-25 14:30:00 CST",
                "2026-11-25 14:30:00 CET",
                "2026-11-25 14:30:00 JST",
                "2026-11-25 14:30:00 SGT",
                "2026-11-25 14:30:00 AEST",
                "2026-11-25T14:30:00IST",
                "2026-11-25 14:30:00 (IST)",
                "2026-11-25 14:30:00.123456 IST",
                "2026-11-25 14:30:00 GMT+05:30",
                "2026-11-25 14:30:00 UTC+05:30",
                "25-11-2026 14:30:00 IST",
                "25/11/2026 02:30:00 PM IST",
                "25 Nov 2026, 02:30:00 PM IST",
            ]
            for raw_val in tz_valid:
                self.assertEqual(parse_date_only(raw_val), date(2026, 11, 25), f"Failed for {raw_val!r}")
                self.assertEqual(strip_to_date_str(raw_val), "2026-11-25", f"Failed strip for {raw_val!r}")

            reject_tokens = [
                "1-2",
                "10-20",
                "21-12",
                "12-34",
                "1/2",
                "10/20",
                "21/12",
                "12/34",
                "2026",
                "2026-11",
                "11-2026",
                "2026/11",
                "11/2026",
                "Nov 2026",
                "Nov 2026 14:30",
                "2026.11",
                "10.5",
                "2:04",
                "14:30",
                "14:30:00",
                "14:30:00 IST",
                "2026-13-01",
                "2026-02-30",
                "0001-01-01",
                "9999-12-31",
                "1969-12-31",
                "2101-01-01",
                "2026-11-25 XYZ",
                "2026-11-25 14:30:00 XYZ",
            ]
            for bad in reject_tokens:
                self.assertIsNone(parse_date_only(bad), f"Expected None for {bad!r}")
                self.assertEqual(strip_to_date_str(bad), "", f"Expected '' for {bad!r}")

            # Non-regression checks
            self.assertEqual(parse_date_only("2026-02-01"), date(2026, 2, 1))
            self.assertEqual(parse_date_only("2026-1-5"), date(2026, 1, 5))
            self.assertEqual(parse_date_only("01-02-2026"), date(2026, 2, 1))
            self.assertEqual(parse_date_only("01/12/2026"), date(2026, 12, 1))
            self.assertEqual(parse_date_only("11/25/2026"), date(2026, 11, 25))
            self.assertEqual(parse_date_only("25 Nov 2026"), date(2026, 11, 25))
            self.assertEqual(parse_date_only("20261125"), date(2026, 11, 25))
            self.assertEqual(parse_date_only(46351), date(2026, 11, 25))
            self.assertEqual(parse_date_only("46351.99"), date(2026, 11, 25))

        user_warnings = [w for w in caught if issubclass(w.category, UserWarning)]
        self.assertEqual(user_warnings, [], f"Unexpected UserWarnings: {[str(w.message) for w in user_warnings]}")

    def test_exporter_float_and_infinity_hardening(self):
        df = pd.DataFrame(
            {
                "Hold total time(Days)": [10.5, float("inf"), float("-inf"), 10.0, 0.0, -18.0, np.nan],
                "Approval Time taken (Days)": pd.Series([1, pd.NA, -18, 0, 2, 3, 4], dtype="Int64"),
                "NonDelta": [1.5, float("inf"), float("-inf"), np.nan, 0.0, np.bool_(True), "ok"],
            }
        )

        csv_text = export_oms_date_only_csv(df)
        rows = list(csv.DictReader(io.StringIO(csv_text)))
        self.assertEqual(
            [r["Hold total time(Days)"] for r in rows],
            ["10.5", "", "", "10", "0", "-18", ""],
        )
        self.assertEqual(
            [r["Approval Time taken (Days)"] for r in rows],
            ["1", "", "-18", "0", "2", "3", "4"],
        )

        xlsx_bytes = export_oms_date_only_excel(df)
        wb = openpyxl.load_workbook(io.BytesIO(xlsx_bytes))
        ws = wb.active
        sheet_rows = list(ws.iter_rows(values_only=True))
        hold_col = [row[0] for row in sheet_rows[1:]]
        self.assertEqual(hold_col, [10.5, None, None, 10, 0, -18, None])
        self.assertIs(type(hold_col[3]), int)
        self.assertIs(type(hold_col[4]), int)
        self.assertIs(type(hold_col[5]), int)

        non_delta_col = [row[2] for row in sheet_rows[1:]]
        self.assertIsNone(non_delta_col[1])
        self.assertIsNone(non_delta_col[2])
        self.assertIs(non_delta_col[5], True)

        # Formula-hint string in delta column and 0-row DataFrame
        hint_df = pd.DataFrame({"Hold total time(Days)": ["colomn21-colomn19", None]})
        hint_csv = export_oms_date_only_csv(hint_df)
        self.assertEqual([r["Hold total time(Days)"] for r in csv.DictReader(io.StringIO(hint_csv))], ["", ""])
        hint_xlsx = export_oms_date_only_excel(hint_df)
        ws_hint = openpyxl.load_workbook(io.BytesIO(hint_xlsx)).active
        self.assertEqual(ws_hint.cell(row=2, column=1).value, "colomn21-colomn19")

        empty_df = pd.DataFrame(columns=["Ticket ID", "Hold total time(Days)"])
        empty_csv = export_oms_date_only_csv(empty_df)
        self.assertEqual(empty_csv.strip(), "Ticket ID,Hold total time(Days)")
        empty_xlsx = export_oms_date_only_excel(empty_df)
        self.assertGreater(len(empty_xlsx), 0)

    def test_convert_coexisting_resolved_total_time_columns(self):
        df = pd.DataFrame(
            [
                {
                    "Ticket ID": "EX-COEXIST-1",
                    "Assigned At": "2026-11-10 08:00:00",
                    "Resolved At Time": "2026-11-12 14:30:00",
                    RESOLVED_TIME_RAW_COL: "2:06",
                    RESOLVED_TIME_DATE_ONLY_COL: 999,
                }
            ]
        )
        converted = convert_oms_to_date_only(df, today_date=date(2026, 11, 25))
        self.assertEqual(len(converted.columns), len(set(converted.columns)))
        self.assertIn(RESOLVED_TIME_RAW_COL, converted.columns)
        self.assertIn(RESOLVED_TIME_DATE_ONLY_COL, converted.columns)
        self.assertEqual(str(converted[RESOLVED_TIME_DATE_ONLY_COL].dtype), "Int64")
        self.assertEqual(int(converted.loc[0, RESOLVED_TIME_DATE_ONLY_COL]), 2)

    def test_raw_dump_standardization_to_42_columns(self):
        from core import TARGET_OMS_COLUMNS, EXCLUDED_OMS_RAW_COLUMNS, is_full_raw_oms_dump
        raw_row = {
            "Ticket ID": "EX-2026-99999",
            "Ticket Source": "Exotel",
            "Status": "Assigned",
            "Call Type": "Inbound",
            "Customer Phone": "+919876543210",
            "Call SID": "CA123456789",
            "Issue Category": "Charger",
            "Issue Subcategory": "Fault",
            "Department": "O&M",
            "Asset": "Station Alpha",
            "Station": "Station Alpha",
            "Assigned To": "Engineer One",
            "Created By": "Agent One",
            "Created At": "24 Nov 2026, 06:08:50 PM",
            "Updated At": "25 Nov 2026, 10:13:49 AM",
            "Location": "Mumbai",
            "Priority": "High",
            "Ticket Type": "EXOTEL",
            "Asset ID": "A-101",
            "Charger": "MH0001",
            "Acknowledgement SLA Deadline": "25 Nov 2026, 04:37:51 PM",
            "Resolution SLA Deadline": "26 Nov 2026, 01:59:45 AM",
            "SLA Breached": "No",
            "Closed At": None,
            "Is Reverted": False,
            "Reverted At": None,
            "Issue Title": "Charger Outage",
            "Creator's Department": "CET",
            "Acknowledged On": "25 Nov 2026, 09:00:00 AM",
            "Acknowledgement SLA Status": "Acknowledged",
            "Resolution Note": "Fixed loose connection",
            "Assigned At": "24 Nov 2026, 06:15:00 PM",
            "In Progress Status Time": "25 Nov 2026, 09:30:00 AM",
            "Hold Status Time": "25 Nov 2026, 12:00:00 PM",
            "Resolved At Time": "26 Nov 2026, 02:00:00 PM",
            "Restored At Time": "26 Nov 2026, 01:00:00 PM",
            "Approval Requested At Time": "26 Nov 2026, 02:30:00 PM",
            "Approval Accepted At Time": "26 Nov 2026, 03:00:00 PM",
            "CET Approval Requested At Time": "26 Nov 2026, 03:15:00 PM",
            "CET Approval Accepted At Time": "26 Nov 2026, 04:00:00 PM",
        }
        df_raw = pd.DataFrame([raw_row])
        self.assertTrue(is_full_raw_oms_dump(df_raw))

        converted = convert_oms_to_date_only(df_raw, today_date=date(2026, 11, 27))

        self.assertEqual(len(converted.columns), 43)
        self.assertEqual(list(converted.columns), list(TARGET_OMS_COLUMNS))

        for excluded in EXCLUDED_OMS_RAW_COLUMNS:
            self.assertNotIn(excluded, converted.columns)

        self.assertEqual(converted.loc[0, "Ticket ID"], "EX-2026-99999")
        self.assertEqual(converted.loc[0, "Creator's Department"], "CET")
        self.assertEqual(converted.loc[0, "SLA Breached As per OMS"], "No")
        self.assertEqual(converted.loc[0, "Created At"], "2026-11-24")
        self.assertEqual(converted.loc[0, "Updated At"], "2026-11-25")
        self.assertEqual(int(converted.loc[0, "Acknowledgement time taken(Days)"]), 1)
        self.assertEqual(int(converted.loc[0, "In Progress total time(Days)"]), 0)
        self.assertEqual(int(converted.loc[0, "Hold total time(Days)"]), 2)
        self.assertEqual(int(converted.loc[0, "Resolved  total time(Days)"]), 2)
        self.assertEqual(int(converted.loc[0, "Restore total time(Days)"]), 2)
        self.assertEqual(int(converted.loc[0, "SLA Breached As per Actual (Days)"]), 2)
        self.assertEqual(int(converted.loc[0, "Approval Time taken (Days)"]), 0)
        self.assertEqual(int(converted.loc[0, "CET Approval Time taken (Days)"]), 0)

        # Excel export check
        xlsx_bytes = export_oms_date_only_excel(converted)
        wb = openpyxl.load_workbook(io.BytesIO(xlsx_bytes))
        self.assertEqual(wb.active.title, "Tickets")
        ws = wb.active
        exported_headers = [ws.cell(row=1, column=c).value for c in range(1, ws.max_column + 1)]
        self.assertEqual(exported_headers, list(TARGET_OMS_COLUMNS))



if __name__ == "__main__":
    unittest.main()
