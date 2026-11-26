"""Deterministic opaque-box fixtures and multi-format serializers for E2E tests."""

from __future__ import annotations

import csv
import html
import io
from datetime import date, datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple
import xml.etree.ElementTree as ET

import openpyxl
import pandas as pd

PROJECT_ROOT = Path(__file__).resolve().parent.parent
DISK_FIR_PATH = Path(r"D:\Comperator\DUMP\FIR DUMP.xlsx")
DISK_OMS_PATH = Path(r"D:\Comperator\DUMP\OMS DUMP.xlsx")

REFERENCE_TODAY = date(2026, 11, 25)
ALT_TODAY = date(2026, 12, 1)

RAW_RESOLVED_COL = "Resolved  total time(Days:Hours)"
DATE_ONLY_RESOLVED_COL = "Resolved  total time(Days)"

DATETIME_COLS_15: Tuple[str, ...] = (
    "Updated At",
    "Created At",
    "Assigned At",
    "Acknowledgement SLA Deadline",
    "Acknowledged On",
    "In Progress Status Time",
    "Hold Status Time",
    "Resolved At Time",
    "Restored At Time",
    "Closed At",
    "Resolution SLA Deadline",
    "Approval Requested At Time",
    "Approval Accepted At Time",
    "CET Approval Requested At Time",
    "CET Approval Accepted At Time",
)

DELTA_COLS_8: Tuple[str, ...] = (
    "Acknowledgement time taken(Days)",
    "Hold total time(Days)",
    "Resolved  total time(Days)",
    "Restore total time(Days)",
    "Total time taken (Days)",
    "SLA Breached As per Actual (Days)",
    "Approval Time taken (Days)",
    "CET Approval Time taken (Days)",
)

RAW_FORMULA_HINTS: Dict[str, str] = {
    "Acknowledgement time taken(Days)": "colomn21-colomn19",
    "Hold total time(Days)": "today-colomn25",
    RAW_RESOLVED_COL: "colomn27-colomn19",
    "Restore total time(Days)": "colomn29-colomn19",
    "Total time taken (Days)": "colomn31-colomn19",
    "SLA Breached As per Actual (Days)": "colomn33-colomn19",
    "Approval Time taken (Days)": "colomn37-colomn36",
    "CET Approval Time taken (Days)": "colomn40-colomn39",
}

REFERENCE_42_COLUMNS_RAW: Tuple[str, ...] = (
    "Sno",                                  # 1
    "Ticket ID",                            # 2
    "Issue Category",                       # 3
    "Issue Subcategory",                    # 4
    "Station",                              # 5
    "Assigned To",                          # 6
    "Zone",                                 # 7
    "ZC-Approval",                          # 8
    "Created By",                           # 9
    "Created At",                           # 10 (Date #2)
    "Updated At",                           # 11 (Date #1)
    "Location",                             # 12
    "Priority",                             # 13
    "Status",                               # 14
    "Ticket Type",                          # 15
    "Department",                           # 16
    "Charger",                              # 17
    "Created Month",                        # 18
    "Assigned At",                          # 19 (colomn19, Date #3)
    "Acknowledgement SLA Deadline",         # 20 (Date #4)
    "Acknowledged On",                      # 21 (colomn21, Date #5)
    "Acknowledgement time taken(Days)",     # 22 (Delta #1)
    "In Progress Status Time",              # 23 (Date #6)
    "Hold Reason",                          # 24
    "Hold Status Time",                     # 25 (colomn25, Date #7)
    "Hold total time(Days)",                # 26 (Delta #2)
    "Resolved At Time",                     # 27 (colomn27, Date #8)
    RAW_RESOLVED_COL,                       # 28 (Delta #3 - double space)
    "Restored At Time",                     # 29 (colomn29, Date #9)
    "Restore total time(Days)",             # 30 (Delta #4)
    "Closed At",                            # 31 (colomn31, Date #10)
    "Total time taken (Days)",              # 32 (Delta #5)
    "Resolution SLA Deadline",              # 33 (colomn33, Date #11)
    "SLA Breached",                         # 34
    "SLA Breached As per Actual (Days)",    # 35 (Delta #6)
    "Approval Requested At Time",           # 36 (colomn36, Date #12)
    "Approval Accepted At Time",            # 37 (colomn37, Date #13)
    "Approval Time taken (Days)",           # 38 (Delta #7)
    "CET Approval Requested At Time",       # 39 (colomn39, Date #14)
    "CET Approval Accepted At Time",        # 40 (colomn40, Date #15)
    "CET Approval Time taken (Days)",       # 41 (Delta #8)
    "Remarks",                              # 42
)

REFERENCE_42_COLUMNS_DATE_ONLY: Tuple[str, ...] = tuple(
    DATE_ONLY_RESOLVED_COL if c == RAW_RESOLVED_COL else c
    for c in REFERENCE_42_COLUMNS_RAW
)

REFERENCE_TICKET_IDS: Tuple[str, ...] = (
    "EX-2026-49803",
    "EX-2026-42245",
    "EX-2026-46397",
    "IT-2026-49625",
    "EX-2026-40520",
    "EX-2026-32559",
    "EX-2026-34312",
    "EX-2026-46325",
    "EX-2026-43053",
)

# Expected 8 delta columns when "today" = 2026-11-25
EXPECTED_DELTAS_2026_11_25: Dict[str, Dict[str, Optional[int]]] = {
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

EXPECTED_HOLD_DELTAS_2026_12_01: Dict[str, Optional[int]] = {
    "EX-2026-49803": None,
    "EX-2026-42245": None,
    "EX-2026-46397": 6,
    "IT-2026-49625": None,
    "EX-2026-40520": 16,
    "EX-2026-32559": 25,
    "EX-2026-34312": None,
    "EX-2026-46325": None,
    "EX-2026-43053": None,
}

# Expected YYYY-MM-DD stripped dates for all 15 date columns across the 9 reference rows
EXPECTED_STRIPPED_DATES: Dict[str, Dict[str, str]] = {
    "EX-2026-49803": {
        "Created At": "2026-11-25",
        "Updated At": "2026-11-25",
        "Assigned At": "",
        "Acknowledgement SLA Deadline": "2026-11-25",
        "Acknowledged On": "",
        "In Progress Status Time": "",
        "Hold Status Time": "",
        "Resolved At Time": "",
        "Restored At Time": "",
        "Closed At": "",
        "Resolution SLA Deadline": "",
        "Approval Requested At Time": "",
        "Approval Accepted At Time": "",
        "CET Approval Requested At Time": "",
        "CET Approval Accepted At Time": "",
    },
    "EX-2026-42245": {
        "Created At": "2026-11-10",
        "Updated At": "2026-11-19",
        "Assigned At": "2026-11-10",
        "Acknowledgement SLA Deadline": "2026-11-11",
        "Acknowledged On": "2026-11-10",
        "In Progress Status Time": "2026-11-11",
        "Hold Status Time": "",
        "Resolved At Time": "",
        "Restored At Time": "2026-11-19",
        "Closed At": "",
        "Resolution SLA Deadline": "2026-11-11",
        "Approval Requested At Time": "",
        "Approval Accepted At Time": "",
        "CET Approval Requested At Time": "",
        "CET Approval Accepted At Time": "",
    },
    "EX-2026-46397": {
        "Created At": "2026-11-21",
        "Updated At": "2026-11-25",
        "Assigned At": "2026-11-21",
        "Acknowledgement SLA Deadline": "2026-11-22",
        "Acknowledged On": "2026-11-23",
        "In Progress Status Time": "2026-11-23",
        "Hold Status Time": "2026-11-25",
        "Resolved At Time": "",
        "Restored At Time": "",
        "Closed At": "",
        "Resolution SLA Deadline": "2026-11-25",
        "Approval Requested At Time": "",
        "Approval Accepted At Time": "",
        "CET Approval Requested At Time": "",
        "CET Approval Accepted At Time": "",
    },
    "IT-2026-49625": {
        "Created At": "2026-11-22",
        "Updated At": "2026-11-23",
        "Assigned At": "2026-11-22",
        "Acknowledgement SLA Deadline": "2026-11-23",
        "Acknowledged On": "2026-11-23",
        "In Progress Status Time": "2026-11-23",
        "Hold Status Time": "",
        "Resolved At Time": "2026-11-23",
        "Restored At Time": "",
        "Closed At": "",
        "Resolution SLA Deadline": "2026-11-24",
        "Approval Requested At Time": "",
        "Approval Accepted At Time": "",
        "CET Approval Requested At Time": "",
        "CET Approval Accepted At Time": "",
    },
    "EX-2026-40520": {
        "Created At": "2026-11-12",
        "Updated At": "2026-11-22",
        "Assigned At": "2026-11-12",
        "Acknowledgement SLA Deadline": "2026-11-12",
        "Acknowledged On": "2026-11-12",
        "In Progress Status Time": "2026-11-13",
        "Hold Status Time": "2026-11-15",
        "Resolved At Time": "",
        "Restored At Time": "2026-11-22",
        "Closed At": "",
        "Resolution SLA Deadline": "2026-11-23",
        "Approval Requested At Time": "2026-11-14",
        "Approval Accepted At Time": "2026-11-15",
        "CET Approval Requested At Time": "",
        "CET Approval Accepted At Time": "",
    },
    "EX-2026-32559": {
        "Created At": "2026-11-03",
        "Updated At": "2026-11-24",
        "Assigned At": "2026-11-03",
        "Acknowledgement SLA Deadline": "2026-11-03",
        "Acknowledged On": "2026-11-03",
        "In Progress Status Time": "2026-11-04",
        "Hold Status Time": "2026-11-06",
        "Resolved At Time": "2026-11-21",
        "Restored At Time": "2026-11-22",
        "Closed At": "",
        "Resolution SLA Deadline": "2026-11-05",
        "Approval Requested At Time": "2026-11-23",
        "Approval Accepted At Time": "2026-11-05",
        "CET Approval Requested At Time": "2026-11-07",
        "CET Approval Accepted At Time": "2026-11-24",
    },
    "EX-2026-34312": {
        "Created At": "2026-11-04",
        "Updated At": "2026-11-20",
        "Assigned At": "2026-11-04",
        "Acknowledgement SLA Deadline": "2026-11-05",
        "Acknowledged On": "2026-11-07",
        "In Progress Status Time": "2026-11-07",
        "Hold Status Time": "",
        "Resolved At Time": "",
        "Restored At Time": "2026-11-20",
        "Closed At": "2026-11-04",
        "Resolution SLA Deadline": "2026-11-10",
        "Approval Requested At Time": "2026-11-08",
        "Approval Accepted At Time": "2026-11-08",
        "CET Approval Requested At Time": "",
        "CET Approval Accepted At Time": "",
    },
    "EX-2026-46325": {
        "Created At": "2026-11-18",
        "Updated At": "2026-11-21",
        "Assigned At": "2026-11-18",
        "Acknowledgement SLA Deadline": "2026-11-19",
        "Acknowledged On": "2026-11-19",
        "In Progress Status Time": "2026-11-19",
        "Hold Status Time": "",
        "Resolved At Time": "2026-11-21",
        "Restored At Time": "",
        "Closed At": "",
        "Resolution SLA Deadline": "2026-11-22",
        "Approval Requested At Time": "2026-11-19",
        "Approval Accepted At Time": "2026-11-20",
        "CET Approval Requested At Time": "2026-11-20",
        "CET Approval Accepted At Time": "2026-11-21",
    },
    "EX-2026-43053": {
        "Created At": "2026-11-14",
        "Updated At": "2026-11-21",
        "Assigned At": "2026-11-14",
        "Acknowledgement SLA Deadline": "2026-11-15",
        "Acknowledged On": "2026-11-21",
        "In Progress Status Time": "",
        "Hold Status Time": "",
        "Resolved At Time": "",
        "Restored At Time": "",
        "Closed At": "",
        "Resolution SLA Deadline": "2026-11-23",
        "Approval Requested At Time": "",
        "Approval Accepted At Time": "",
        "CET Approval Requested At Time": "",
        "CET Approval Accepted At Time": "",
    },
}


_RAW_TIMESTAMP_ROWS: List[Dict[str, Any]] = [
    {
        "Sno": 1,
        "Ticket ID": "EX-2026-49803",
        "Issue Category": "Hardware",
        "Issue Subcategory": "Connector Fault",
        "Station": "Alpha Hub West",
        "Assigned To": "Unassigned",
        "Zone": "West",
        "ZC-Approval": "Pending",
        "Created By": "NOC Bot",
        "Created At": "2026-11-25 09:15:00",
        "Updated At": "2026-11-25 09:20:00",
        "Location": "Mumbai",
        "Priority": "High",
        "Status": "Open",
        "Ticket Type": "External",
        "Department": "O&M",
        "Charger": "MH0011",
        "Created Month": "Nov-2026",
        "Assigned At": "",
        "Acknowledgement SLA Deadline": "2026-11-25 13:15:00",
        "Acknowledged On": "",
        "Acknowledgement time taken(Days)": "colomn21-colomn19",
        "In Progress Status Time": "",
        "Hold Reason": "",
        "Hold Status Time": "",
        "Hold total time(Days)": "today-colomn25",
        "Resolved At Time": "",
        RAW_RESOLVED_COL: "colomn27-colomn19",
        "Restored At Time": "",
        "Restore total time(Days)": "colomn29-colomn19",
        "Closed At": "",
        "Total time taken (Days)": "colomn31-colomn19",
        "Resolution SLA Deadline": "",
        "SLA Breached": "No",
        "SLA Breached As per Actual (Days)": "colomn33-colomn19",
        "Approval Requested At Time": "",
        "Approval Accepted At Time": "",
        "Approval Time taken (Days)": "colomn37-colomn36",
        "CET Approval Requested At Time": "",
        "CET Approval Accepted At Time": "",
        "CET Approval Time taken (Days)": "colomn40-colomn39",
        "Remarks": "Awaiting field assignment",
    },
    {
        "Sno": 2,
        "Ticket ID": "EX-2026-42245",
        "Issue Category": "Power",
        "Issue Subcategory": "Discom_Issue",
        "Station": "Delhi Hub North",
        "Assigned To": "A. Sharma",
        "Zone": "North",
        "ZC-Approval": "Not Required",
        "Created By": "Ops Desk",
        "Created At": "2026-11-10 20:00:00",
        "Updated At": "2026-11-19 08:35:00",
        "Location": "Delhi",
        "Priority": "High",
        "Status": "Resolved",
        "Ticket Type": "External",
        "Department": "Field Ops",
        "Charger": "DL0022",
        "Created Month": "Nov-2026",
        "Assigned At": "2026-11-10 22:15:00",
        "Acknowledgement SLA Deadline": "2026-11-11 02:15:00",
        "Acknowledged On": "2026-11-10 23:45:00",
        "Acknowledgement time taken(Days)": "",
        "In Progress Status Time": "2026-11-11 09:00:00",
        "Hold Reason": "",
        "Hold Status Time": "",
        "Hold total time(Days)": "",
        "Resolved At Time": "",
        RAW_RESOLVED_COL: "0 Days: 0 Hours",
        "Restored At Time": "2026-11-19 08:30:00",
        "Restore total time(Days)": "",
        "Closed At": "",
        "Total time taken (Days)": "",
        "Resolution SLA Deadline": "2026-11-11 04:00:00",
        "SLA Breached": "No",
        "SLA Breached As per Actual (Days)": "",
        "Approval Requested At Time": "",
        "Approval Accepted At Time": "",
        "Approval Time taken (Days)": "",
        "CET Approval Requested At Time": "",
        "CET Approval Accepted At Time": "",
        "CET Approval Time taken (Days)": "",
        "Remarks": "Power restored by discom",
    },
    {
        "Sno": 3,
        "Ticket ID": "EX-2026-46397",
        "Issue Category": "Hardware",
        "Issue Subcategory": "Display Module",
        "Station": "Bengaluru Central",
        "Assigned To": "R. Nair",
        "Zone": "South",
        "ZC-Approval": "Approved",
        "Created By": "Ops Desk",
        "Created At": "2026-11-21 17:30:00",
        "Updated At": "2026-11-25 11:25:00",
        "Location": "Bengaluru",
        "Priority": "Critical",
        "Status": "On Hold",
        "Ticket Type": "External",
        "Department": "Maintenance",
        "Charger": "KA0033",
        "Created Month": "Nov-2026",
        "Assigned At": "2026-11-21 18:40:00",
        "Acknowledgement SLA Deadline": "2026-11-22 10:00:00",
        "Acknowledged On": "2026-11-23 07:10:00",
        "Acknowledgement time taken(Days)": "",
        "In Progress Status Time": "2026-11-23 12:00:00",
        "Hold Reason": "Spare part awaited",
        "Hold Status Time": "2026-11-25 11:20:00",
        "Hold total time(Days)": "",
        "Resolved At Time": "",
        RAW_RESOLVED_COL: "",
        "Restored At Time": "",
        "Restore total time(Days)": "",
        "Closed At": "",
        "Total time taken (Days)": "",
        "Resolution SLA Deadline": "2026-11-25 09:00:00",
        "SLA Breached": "Yes",
        "SLA Breached As per Actual (Days)": "",
        "Approval Requested At Time": "",
        "Approval Accepted At Time": "",
        "Approval Time taken (Days)": "",
        "CET Approval Requested At Time": "",
        "CET Approval Accepted At Time": "",
        "CET Approval Time taken (Days)": "",
        "Remarks": "Waiting OEM dispatch",
    },
    {
        "Sno": 4,
        "Ticket ID": "IT-2026-49625",
        "Issue Category": "Network",
        "Issue Subcategory": "OCPP Offline",
        "Station": "Kolkata East Park",
        "Assigned To": "S. Das",
        "Zone": "East",
        "ZC-Approval": "Not Required",
        "Created By": "NOC Bot",
        "Created At": "2026-11-22 19:00:00",
        "Updated At": "2026-11-23 20:00:00",
        "Location": "Kolkata",
        "Priority": "Medium",
        "Status": "Resolved",
        "Ticket Type": "Internal",
        "Department": "NOC",
        "Charger": "WB0044",
        "Created Month": "Nov-2026",
        "Assigned At": "2026-11-22 21:00:00",
        "Acknowledgement SLA Deadline": "2026-11-23 09:00:00",
        "Acknowledged On": "2026-11-23 02:30:00",
        "Acknowledgement time taken(Days)": "",
        "In Progress Status Time": "2026-11-23 08:00:00",
        "Hold Reason": "",
        "Hold Status Time": "",
        "Hold total time(Days)": "",
        "Resolved At Time": "2026-11-23 19:45:00",
        RAW_RESOLVED_COL: "0 Days: 22 Hours",
        "Restored At Time": "",
        "Restore total time(Days)": "",
        "Closed At": "",
        "Total time taken (Days)": "",
        "Resolution SLA Deadline": "2026-11-24 12:00:00",
        "SLA Breached": "No",
        "SLA Breached As per Actual (Days)": "",
        "Approval Requested At Time": "",
        "Approval Accepted At Time": "",
        "Approval Time taken (Days)": "",
        "CET Approval Requested At Time": "",
        "CET Approval Accepted At Time": "",
        "CET Approval Time taken (Days)": "",
        "Remarks": "Modem rebooted",
    },
    {
        "Sno": 5,
        "Ticket ID": "EX-2026-40520",
        "Issue Category": "Civil",
        "Issue Subcategory": "Earthing Pit",
        "Station": "Pune West Plaza",
        "Assigned To": "M. Patil",
        "Zone": "West",
        "ZC-Approval": "Approved",
        "Created By": "Ops Desk",
        "Created At": "2026-11-12 09:30:00",
        "Updated At": "2026-11-22 09:10:00",
        "Location": "Pune",
        "Priority": "High",
        "Status": "On Hold",
        "Ticket Type": "External",
        "Department": "O&M",
        "Charger": "MH0055",
        "Created Month": "Nov-2026",
        "Assigned At": "2026-11-12 10:05:00",
        "Acknowledgement SLA Deadline": "2026-11-12 18:00:00",
        "Acknowledged On": "2026-11-12 16:20:00",
        "Acknowledgement time taken(Days)": "",
        "In Progress Status Time": "2026-11-13 10:00:00",
        "Hold Reason": "Site access permit",
        "Hold Status Time": "2026-11-15 14:10:00",
        "Hold total time(Days)": "",
        "Resolved At Time": "",
        RAW_RESOLVED_COL: "",
        "Restored At Time": "2026-11-22 09:00:00",
        "Restore total time(Days)": "",
        "Closed At": "",
        "Total time taken (Days)": "",
        "Resolution SLA Deadline": "2026-11-23 10:05:00",
        "SLA Breached": "No",
        "SLA Breached As per Actual (Days)": "",
        "Approval Requested At Time": "2026-11-14 23:30:00",
        "Approval Accepted At Time": "2026-11-15 08:15:00",
        "Approval Time taken (Days)": "",
        "CET Approval Requested At Time": "",
        "CET Approval Accepted At Time": "",
        "CET Approval Time taken (Days)": "",
        "Remarks": "Permit cleared",
    },
    {
        "Sno": 6,
        "Ticket ID": "EX-2026-32559",
        "Issue Category": "Hardware",
        "Issue Subcategory": "Rectifier Module",
        "Station": "Ahmedabad Depot",
        "Assigned To": "K. Mehta",
        "Zone": "West",
        "ZC-Approval": "Pending CET",
        "Created By": "Ops Desk",
        "Created At": "2026-11-03 08:15:00",
        "Updated At": "2026-11-24 12:35:00",
        "Location": "Ahmedabad",
        "Priority": "Critical",
        "Status": "Pending CET Approval",
        "Ticket Type": "External",
        "Department": "O&M",
        "Charger": "GJ0066",
        "Created Month": "Nov-2026",
        "Assigned At": "2026-11-03 09:00:00",
        "Acknowledgement SLA Deadline": "2026-11-03 17:00:00",
        "Acknowledged On": "2026-11-03 17:45:00",
        "Acknowledgement time taken(Days)": "",
        "In Progress Status Time": "2026-11-04 10:00:00",
        "Hold Reason": "CET signoff",
        "Hold Status Time": "2026-11-06 11:30:00",
        "Hold total time(Days)": "",
        "Resolved At Time": "2026-11-21 15:20:00",
        RAW_RESOLVED_COL: "18 Days: 6 Hours",
        "Restored At Time": "2026-11-22 10:00:00",
        "Restore total time(Days)": "",
        "Closed At": "",
        "Total time taken (Days)": "",
        "Resolution SLA Deadline": "2026-11-05 09:00:00",
        "SLA Breached": "Yes",
        "SLA Breached As per Actual (Days)": "",
        "Approval Requested At Time": "2026-11-23 14:00:00",
        "Approval Accepted At Time": "2026-11-05 16:00:00",
        "Approval Time taken (Days)": "",
        "CET Approval Requested At Time": "2026-11-07 10:00:00",
        "CET Approval Accepted At Time": "2026-11-24 12:30:00",
        "CET Approval Time taken (Days)": "",
        "Remarks": "Backdated approval entry",
    },
    {
        "Sno": 7,
        "Ticket ID": "EX-2026-34312",
        "Issue Category": "Software",
        "Issue Subcategory": "Firmware Update",
        "Station": "Chennai South Hub",
        "Assigned To": "V. Iyer",
        "Zone": "South",
        "ZC-Approval": "Approved",
        "Created By": "Field Tech",
        "Created At": "2026-11-04 11:00:00",
        "Updated At": "2026-11-20 18:05:00",
        "Location": "Chennai",
        "Priority": "Low",
        "Status": "Closed",
        "Ticket Type": "External",
        "Department": "Engineering",
        "Charger": "TN0077",
        "Created Month": "Nov-2026",
        "Assigned At": "2026-11-04 13:15:00",
        "Acknowledgement SLA Deadline": "2026-11-05 13:15:00",
        "Acknowledged On": "2026-11-07 09:40:00",
        "Acknowledgement time taken(Days)": "",
        "In Progress Status Time": "2026-11-07 14:00:00",
        "Hold Reason": "",
        "Hold Status Time": "",
        "Hold total time(Days)": "",
        "Resolved At Time": "",
        RAW_RESOLVED_COL: "",
        "Restored At Time": "2026-11-20 18:00:00",
        "Restore total time(Days)": "",
        "Closed At": "2026-11-04 20:00:00",
        "Total time taken (Days)": "",
        "Resolution SLA Deadline": "2026-11-10 13:15:00",
        "SLA Breached": "No",
        "SLA Breached As per Actual (Days)": "",
        "Approval Requested At Time": "2026-11-08 08:00:00",
        "Approval Accepted At Time": "2026-11-08 19:30:00",
        "Approval Time taken (Days)": "",
        "CET Approval Requested At Time": "",
        "CET Approval Accepted At Time": "",
        "CET Approval Time taken (Days)": "",
        "Remarks": "Same-day closure",
    },
    {
        "Sno": 8,
        "Ticket ID": "EX-2026-46325",
        "Issue Category": "Hardware",
        "Issue Subcategory": "Cable Replacement",
        "Station": "Jaipur Hub North",
        "Assigned To": "P. Verma",
        "Zone": "North",
        "ZC-Approval": "Pending",
        "Created By": "Ops Desk",
        "Created At": "2026-11-18 09:45:00",
        "Updated At": "2026-11-21 17:00:00",
        "Location": "Jaipur",
        "Priority": "Medium",
        "Status": "Pending Approval",
        "Ticket Type": "External",
        "Department": "O&M",
        "Charger": "RJ0088",
        "Created Month": "Nov-2026",
        "Assigned At": "2026-11-18 11:00:00",
        "Acknowledgement SLA Deadline": "2026-11-19 11:00:00",
        "Acknowledged On": "2026-11-19 10:00:00",
        "Acknowledgement time taken(Days)": "",
        "In Progress Status Time": "2026-11-19 11:30:00",
        "Hold Reason": "",
        "Hold Status Time": "",
        "Hold total time(Days)": "",
        "Resolved At Time": "2026-11-21 16:45:00",
        RAW_RESOLVED_COL: "3 Days: 5 Hours",
        "Restored At Time": "",
        "Restore total time(Days)": "",
        "Closed At": "",
        "Total time taken (Days)": "",
        "Resolution SLA Deadline": "2026-11-22 11:00:00",
        "SLA Breached": "No",
        "SLA Breached As per Actual (Days)": "",
        "Approval Requested At Time": "2026-11-19 12:00:00",
        "Approval Accepted At Time": "2026-11-20 09:30:00",
        "Approval Time taken (Days)": "",
        "CET Approval Requested At Time": "2026-11-20 14:00:00",
        "CET Approval Accepted At Time": "2026-11-21 08:15:00",
        "CET Approval Time taken (Days)": "",
        "Remarks": "Awaiting commercial signoff",
    },
    {
        "Sno": 9,
        "Ticket ID": "EX-2026-43053",
        "Issue Category": "Network",
        "Issue Subcategory": "SIM Fault",
        "Station": "Patna East Hub",
        "Assigned To": "R. Singh",
        "Zone": "East",
        "ZC-Approval": "Not Required",
        "Created By": "NOC Bot",
        "Created At": "2026-11-14 06:50:00",
        "Updated At": "2026-11-21 12:15:00",
        "Location": "Patna",
        "Priority": "Low",
        "Status": "Assigned",
        "Ticket Type": "External",
        "Department": "Field Ops",
        "Charger": "BR0099",
        "Created Month": "Nov-2026",
        "Assigned At": "2026-11-14 07:30:00",
        "Acknowledgement SLA Deadline": "2026-11-15 07:30:00",
        "Acknowledged On": "2026-11-21 12:00:00",
        "Acknowledgement time taken(Days)": "",
        "In Progress Status Time": "",
        "Hold Reason": "",
        "Hold Status Time": "",
        "Hold total time(Days)": "",
        "Resolved At Time": "",
        RAW_RESOLVED_COL: "",
        "Restored At Time": "",
        "Restore total time(Days)": "",
        "Closed At": "",
        "Total time taken (Days)": "",
        "Resolution SLA Deadline": "2026-11-23 07:30:00",
        "SLA Breached": "No",
        "SLA Breached As per Actual (Days)": "",
        "Approval Requested At Time": "",
        "Approval Accepted At Time": "",
        "Approval Time taken (Days)": "",
        "CET Approval Requested At Time": "",
        "CET Approval Accepted At Time": "",
        "CET Approval Time taken (Days)": "",
        "Remarks": "Technician scheduled",
    },
]


def _format_mixed_timestamp(iso_ts: str, style_idx: int) -> Any:
    if not iso_ts:
        return ""
    dt = datetime.strptime(iso_ts, "%Y-%m-%d %H:%M:%S")
    mode = style_idx % 5
    if mode == 0:
        return dt.strftime("%Y-%m-%d %H:%M:%S")
    if mode == 1:
        return dt.strftime("%Y-%m-%dT%H:%M:%SZ")
    if mode == 2:
        return dt.strftime("%d-%m-%Y %H:%M")
    if mode == 3:
        return dt.strftime("%d/%m/%Y %H:%M:%S")
    return dt.strftime("%d %b %Y, %I:%M:%S %p")


def build_reference_9_rows_df(date_style: str = "iso") -> pd.DataFrame:
    """Return the exact 9-row x 42-col Nov1-25-OMS raw fixture DataFrame."""
    rows: List[Dict[str, Any]] = []
    for r_idx, base in enumerate(_RAW_TIMESTAMP_ROWS):
        row = dict(base)
        if date_style == "mixed":
            for c_idx, col in enumerate(DATETIME_COLS_15):
                if row.get(col):
                    row[col] = _format_mixed_timestamp(str(row[col]), r_idx + c_idx)
        elif date_style == "datetime":
            for col in DATETIME_COLS_15:
                if row.get(col):
                    row[col] = datetime.strptime(str(row[col]), "%Y-%m-%d %H:%M:%S")
        rows.append(row)
    return pd.DataFrame(rows, columns=list(REFERENCE_42_COLUMNS_RAW))


def build_minimal_fir_df(chargers: Optional[Sequence[Dict[str, str]]] = None) -> pd.DataFrame:
    """Build a valid FIR DataFrame for OMS-FIR Compare tests."""
    if chargers is None:
        chargers = [
            {
                "OCPP ID": "MH0011",
                "Charger Status": "Unavailable",
                "Tenant": "ChargeZone - B2C",
                "Priority": "2.High (1-3 Days)",
                "Zone": "West",
                "OEM": "Ador",
                "Station Name": "Alpha Hub West",
                "City": "Mumbai",
                "Issue Age": "3 Days",
            },
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
            },
        ]
    return pd.DataFrame(list(chargers))


def to_csv_bytes(df: pd.DataFrame, bom: bool = False) -> bytes:
    text = df.to_csv(index=False)
    encoding = "utf-8-sig" if bom else "utf-8"
    return text.encode(encoding)


def to_xlsx_bytes(
    df: pd.DataFrame,
    excel_serial_cols: Sequence[str] = (),
    pre_header_rows: int = 0,
    header_overrides: Optional[Dict[str, str]] = None,
) -> bytes:
    """Serialize DataFrame to valid .xlsx bytes using openpyxl."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "OMS"

    row_cursor = 1
    for idx in range(pre_header_rows):
        ws.cell(row=row_cursor, column=1, value=f"Report Generated Banner Line {idx + 1} @ 25-11-2026")
        row_cursor += 1

    overrides = header_overrides or {}
    columns = list(df.columns)
    for col_idx, col_name in enumerate(columns, start=1):
        ws.cell(row=row_cursor, column=col_idx, value=overrides.get(col_name, col_name))
    row_cursor += 1

    serial_set = set(excel_serial_cols)
    excel_epoch = datetime(1899, 12, 30)

    for _, row in df.iterrows():
        for col_idx, col_name in enumerate(columns, start=1):
            val = row[col_name]
            if pd.isna(val) or val == "":
                cell_val = None
            elif col_name in serial_set:
                if isinstance(val, datetime):
                    dt = val
                else:
                    dt = datetime.strptime(str(val)[:19], "%Y-%m-%d %H:%M:%S")
                cell_val = (dt - excel_epoch).total_seconds() / 86400.0
            else:
                cell_val = val
            ws.cell(row=row_cursor, column=col_idx, value=cell_val)
        row_cursor += 1

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def to_xls_xml_bytes(df: pd.DataFrame, sheet_name: str = "OMS_Dump") -> bytes:
    """Generate a valid Excel 2003 XML (SpreadsheetML) .xls payload without xlwt."""
    ns = "urn:schemas-microsoft-com:office:spreadsheet"
    ET.register_namespace("", ns)
    ET.register_namespace("ss", ns)

    def _q(tag: str) -> str:
        return f"{{{ns}}}{tag}"

    wb_el = ET.Element(_q("Workbook"))
    ws_el = ET.SubElement(wb_el, _q("Worksheet"), {_q("Name"): sheet_name})
    table_el = ET.SubElement(ws_el, _q("Table"))

    # Header row
    hdr_row = ET.SubElement(table_el, _q("Row"))
    for col in df.columns:
        cell = ET.SubElement(hdr_row, _q("Cell"))
        data = ET.SubElement(cell, _q("Data"), {_q("Type"): "String"})
        data.text = str(col)

    # Data rows
    for _, row in df.iterrows():
        r_el = ET.SubElement(table_el, _q("Row"))
        for col in df.columns:
            val = row[col]
            cell = ET.SubElement(r_el, _q("Cell"))
            if pd.isna(val) or val == "":
                data = ET.SubElement(cell, _q("Data"), {_q("Type"): "String"})
                data.text = ""
            elif isinstance(val, (int, float)) and not isinstance(val, bool):
                data = ET.SubElement(cell, _q("Data"), {_q("Type"): "Number"})
                data.text = str(val)
            else:
                data = ET.SubElement(cell, _q("Data"), {_q("Type"): "String"})
                data.text = str(val)

    xml_bytes = ET.tostring(wb_el, encoding="utf-8", xml_declaration=True)
    return xml_bytes


def to_xls_html_bytes(df: pd.DataFrame) -> bytes:
    """Generate an HTML-table .xls payload commonly exported by legacy portals."""
    parts = [
        "<html><head><meta charset='utf-8'></head><body><table>",
        "<tr>" + "".join(f"<th>{html.escape(str(c))}</th>" for c in df.columns) + "</tr>",
    ]
    for _, row in df.iterrows():
        cells = []
        for col in df.columns:
            val = row[col]
            txt = "" if (pd.isna(val) or val == "") else str(val)
            cells.append(f"<td>{html.escape(txt)}</td>")
        parts.append("<tr>" + "".join(cells) + "</tr>")
    parts.append("</table></body></html>")
    return "\n".join(parts).encode("utf-8")


def to_xls_tsv_bytes(df: pd.DataFrame) -> bytes:
    """Generate a tab-delimited .xls payload."""
    buf = io.StringIO()
    writer = csv.writer(buf, delimiter="\t", lineterminator="\n")
    writer.writerow(list(df.columns))
    for _, row in df.iterrows():
        writer.writerow(["" if (pd.isna(row[c]) or row[c] == "") else str(row[c]) for c in df.columns])
    return buf.getvalue().encode("utf-8")


def corrupt_binary_bytes() -> bytes:
    return b"\x00\xff\xfe\x13NOT_AN_EXCEL_OR_CSV_FILE\x89PNG\r\n\x1a\nCORRUPT"
