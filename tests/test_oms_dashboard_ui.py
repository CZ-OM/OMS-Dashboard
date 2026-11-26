import csv
import io
import subprocess
import sys
import unittest
import warnings
from datetime import date
from pathlib import Path

import openpyxl
import pandas as pd
import pyarrow.ipc as pa_ipc
from streamlit.testing.v1 import AppTest

from core import (
    OMS_DATETIME_COLUMNS,
    OMS_DELTA_COLUMNS,
    RESOLVED_TIME_DATE_ONLY_COL,
    RESOLVED_TIME_RAW_COL,
    export_oms_date_only_csv,
    export_oms_date_only_excel,
)

PROJECT_ROOT = Path(__file__).resolve().parent.parent
APP_PATH = PROJECT_ROOT / "app.py"
SAMPLE_OMS_PATH = Path("D:/Comperator/DUMP/OMS DUMP.xlsx")
SAMPLE_FIR_PATH = Path("D:/Comperator/DUMP/FIR DUMP.xlsx")

NOV1_25_OMS_COLUMNS = [
    "Sno",
    "Ticket ID",
    "Issue Category",
    "Issue Subcategory",
    "Station",
    "Assigned To",
    "Zone",
    "ZC-Approval",
    "Created By",
    "Created At",
    "Updated At",
    "Location",
    "Priority",
    "Status",
    "Ticket Type",
    "Department",
    "Charger",
    "Acknowledgement SLA Deadline",
    "Assigned At",
    "Hold Reason",
    "Acknowledged On",
    "Acknowledgement time taken(Days)",
    "In Progress Status Time",
    "SLA Breached",
    "Hold Status Time",
    "Hold total time(Days)",
    "Resolved At Time",
    RESOLVED_TIME_RAW_COL,
    "Restored At Time",
    "Restore total time(Days)",
    "Closed At",
    "Total time taken (Days)",
    "Resolution SLA Deadline",
    "SLA Breached As per Actual (Days)",
    "Approval Status",
    "Approval Requested At Time",
    "Approval Accepted At Time",
    "Approval Time taken (Days)",
    "CET Approval Requested At Time",
    "CET Approval Accepted At Time",
    "CET Approval Time taken (Days)",
    "Created Month",
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
    def _blank_row(sno: int, ticket_id: str, status: str, zone: str = "North", prio: str = "High", sla: str = "No") -> dict:
        row = {col: "" for col in NOV1_25_OMS_COLUMNS}
        row.update(
            {
                "Sno": sno,
                "Ticket ID": ticket_id,
                "Issue Category": "Hardware" if sno % 2 else "Connectivity",
                "Issue Subcategory": "Connector Fault",
                "Station": f"Station {sno}",
                "Assigned To": f"Tech {sno}",
                "Zone": zone,
                "ZC-Approval": "No",
                "Created By": "Ops",
                "Created At": "2026-11-01 09:15:00",
                "Updated At": "2026-11-24 18:30:00",
                "Location": "Delhi",
                "Priority": prio,
                "Status": status,
                "Ticket Type": "Corrective",
                "Department": "Field Ops",
                "Charger": f"CHG-{1000 + sno}",
                "Acknowledgement SLA Deadline": "2026-11-02 09:15:00",
                "SLA Breached": sla,
                "Created Month": "Nov 2026",
            }
        )
        return row

    r1 = _blank_row(1, "EX-2026-49803", "Open", zone="North", prio="Critical", sla="Yes")
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

    r2 = _blank_row(2, "EX-2026-42245", "Restored", zone="South", prio="Medium", sla="No")
    r2.update(
        {
            "Assigned At": "2026-11-10 08:00:00",
            "Acknowledged On": "2026-11-10 14:45:00",
            "Restored At Time": "19-11-2026 11:20:00",
            "Resolution SLA Deadline": "11/11/2026 08:00",
        }
    )

    r3 = _blank_row(3, "EX-2026-46397", "On Hold", zone="East", prio="High", sla="Yes")
    r3.update(
        {
            "Assigned At": "2026-11-20 10:00:00",
            "Acknowledged On": "22 Nov 2026, 09:30:00 AM",
            "Hold Status Time": "2026-11-25 07:10:00",
            "Hold Reason": "Spare parts pending",
            "Resolution SLA Deadline": "2026-11-24 10:00:00",
        }
    )

    r4 = _blank_row(4, "IT-2026-49625", "Resolved", zone="West", prio="Low", sla="No")
    r4.update(
        {
            "Assigned At": "2026-11-18 16:00:00",
            "Acknowledged On": "2026-11-19 09:00:00",
            "Resolved At Time": "2026-11-19 17:30:00",
            "Resolution SLA Deadline": "2026-11-20 16:00:00",
        }
    )

    r5 = _blank_row(5, "EX-2026-40520", "Pending Approval", zone="North", prio="High", sla="No")
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

    r6 = _blank_row(6, "EX-2026-32559", "Pending CET Approval", zone="South", prio="Critical", sla="Yes")
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

    r7 = _blank_row(7, "EX-2026-34312", "Closed", zone="West", prio="Low", sla="No")
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

    r8 = _blank_row(8, "EX-2026-46325", "Resolved", zone="East", prio="Medium", sla="No")
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

    r9 = _blank_row(9, "EX-2026-43053", "Assigned", zone="North", prio="High", sla="No")
    r9.update(
        {
            "Assigned At": "2026-11-10 09:00:00",
            "Acknowledged On": "2026-11-17 09:00:00",
            "Resolution SLA Deadline": "2026-11-19 09:00:00",
        }
    )

    return pd.DataFrame([r1, r2, r3, r4, r5, r6, r7, r8, r9], columns=NOV1_25_OMS_COLUMNS)


def _df_to_xlsx_bytes(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)
    return buf.getvalue()


def _df_to_csv_bytes(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False).encode("utf-8")


class TestOmsDashboardUI(unittest.TestCase):
    def _new_app(self) -> AppTest:
        at = AppTest.from_file(str(APP_PATH), default_timeout=30)
        at.run()
        self.assertEqual(len(at.exception), 0, f"Unexpected startup exceptions: {at.exception}")
        return at

    def _upload_oms_and_set_today(self, at: AppTest, filename: str, raw_bytes: bytes, mime: str, today_val: date) -> None:
        at.sidebar.file_uploader(key="oms_dash_uploader").upload(filename, raw_bytes, mime).run()
        self.assertEqual(len(at.exception), 0, f"Exceptions on upload: {at.exception}")
        at.sidebar.date_input(key="oms_dash_today").set_value(today_val).run()
        self.assertEqual(len(at.exception), 0, f"Exceptions on date_input: {at.exception}")

    def test_default_navigation_and_landing_state(self):
        at = self._new_app()
        nav = at.sidebar.radio(key="nav_view")
        self.assertEqual(list(nav.options), ["OMS Dashboard", "OMS-FIR Compare"])
        self.assertEqual(nav.value, "OMS Dashboard")
        self.assertEqual(at.sidebar.date_input(key="oms_dash_today").value, date.today())
        self.assertIsNone(at.session_state["oms_dash_df"])
        self.assertTrue(
            any("OMS Issue Tracker Dashboard" in m.value for m in at.markdown),
            "Expected OMS landing header on default view",
        )

    def test_upload_reference_oms_xlsx_and_csv_conversion(self):
        ref_df = _build_reference_oms_df()
        cases = [
            (
                "Nov1-25-OMS.xlsx",
                _df_to_xlsx_bytes(ref_df),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            ),
            (
                "Nov1-25-OMS.csv",
                _df_to_csv_bytes(ref_df),
                "text/csv",
            ),
        ]
        for fname, payload, mime in cases:
            with self.subTest(filename=fname):
                at = self._new_app()
                self._upload_oms_and_set_today(at, fname, payload, mime, date(2026, 11, 25))

                conv_df = at.session_state["oms_dash_df"]
                self.assertIsNotNone(conv_df)
                self.assertEqual(len(conv_df), 9)
                self.assertEqual(len(conv_df.columns), 42)
                self.assertIn(RESOLVED_TIME_DATE_ONLY_COL, conv_df.columns)
                self.assertNotIn(RESOLVED_TIME_RAW_COL, conv_df.columns)
                self.assertEqual(conv_df.attrs.get("today_date"), "2026-11-25")

                for dt_col in OMS_DATETIME_COLUMNS:
                    self.assertIn(dt_col, conv_df.columns)
                    for val in conv_df[dt_col].tolist():
                        self.assertTrue(val == "" or (len(val) == 10 and val[4] == "-" and val[7] == "-"))

                by_id = conv_df.set_index("Ticket ID")
                for ticket_id, expected_deltas in EXPECTED_9_ROWS_DELTAS.items():
                    for col_name, expected_val in expected_deltas.items():
                        actual_val = by_id.loc[ticket_id, col_name]
                        if expected_val is None:
                            self.assertTrue(pd.isna(actual_val), f"{ticket_id} {col_name} expected NA, got {actual_val}")
                        else:
                            self.assertEqual(int(actual_val), expected_val, f"{ticket_id} {col_name}")

                self.assertGreaterEqual(len(at.get("vega_lite_chart")), 4)
                tracker_df = next(d.value for d in at.dataframe if "Ticket ID" in d.value.columns and "Status" in d.value.columns)
                schema_df = next(d.value for d in at.dataframe if "No." in d.value.columns and len(d.value.columns) == len(conv_df.columns) + 1)
                self.assertEqual(len(tracker_df), 9)
                self.assertEqual(len(schema_df), 9)
                self.assertEqual(list(tracker_df["No."]), list(range(1, 10)))
                self.assertEqual(list(schema_df.columns), ["No."] + list(conv_df.columns))
                self.assertEqual(len(at.download_button), 4)

    def test_dynamic_today_recalculation_and_clear_fallback(self):
        ref_df = _build_reference_oms_df()
        at = self._new_app()
        self._upload_oms_and_set_today(
            at,
            "Nov1-25-OMS.xlsx",
            _df_to_xlsx_bytes(ref_df),
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            date(2026, 11, 25),
        )

        url_csv_t1 = at.download_button[0].proto.url
        url_xlsx_t1 = at.download_button[1].proto.url

        at.sidebar.date_input(key="oms_dash_today").set_value(date(2026, 12, 5)).run()
        self.assertEqual(len(at.exception), 0)
        conv_t2 = at.session_state["oms_dash_df"]
        self.assertEqual(conv_t2.attrs.get("today_date"), "2026-12-05")
        by_id_t2 = conv_t2.set_index("Ticket ID")
        self.assertEqual(int(by_id_t2.loc["EX-2026-46397", "Hold total time(Days)"]), 10)
        self.assertEqual(int(by_id_t2.loc["EX-2026-40520", "Hold total time(Days)"]), 20)
        self.assertEqual(int(by_id_t2.loc["EX-2026-32559", "Hold total time(Days)"]), 29)
        self.assertTrue(pd.isna(by_id_t2.loc["EX-2026-49803", "Hold total time(Days)"]))

        self.assertNotEqual(at.download_button[0].proto.url, url_csv_t1)
        self.assertNotEqual(at.download_button[1].proto.url, url_xlsx_t1)

        csv_text = export_oms_date_only_csv(conv_t2)
        csv_rows = list(csv.DictReader(io.StringIO(csv_text)))
        csv_by_id = {r["Ticket ID"]: r for r in csv_rows}
        self.assertEqual(csv_by_id["EX-2026-46397"]["Hold total time(Days)"], "10")
        self.assertEqual(csv_by_id["EX-2026-32559"]["Hold total time(Days)"], "29")
        self.assertEqual(csv_by_id["EX-2026-32559"]["Approval Time taken (Days)"], "-18")
        self.assertEqual(csv_by_id["EX-2026-49803"]["Hold total time(Days)"], "")

        xlsx_bytes = export_oms_date_only_excel(conv_t2)
        wb = openpyxl.load_workbook(io.BytesIO(xlsx_bytes))
        ws = wb.active
        headers = [ws.cell(row=1, column=c).value for c in range(1, ws.max_column + 1)]
        self.assertEqual(len(headers), 42)
        self.assertIn(RESOLVED_TIME_DATE_ONLY_COL, headers)
        hold_idx = headers.index("Hold total time(Days)") + 1
        tid_idx = headers.index("Ticket ID") + 1
        wb_by_id = {
            ws.cell(row=r, column=tid_idx).value: ws.cell(row=r, column=hold_idx).value
            for r in range(2, ws.max_row + 1)
        }
        self.assertEqual(wb_by_id["EX-2026-46397"], 10)
        self.assertEqual(wb_by_id["EX-2026-32559"], 29)
        self.assertIsNone(wb_by_id["EX-2026-49803"])

        at.sidebar.date_input(key="oms_dash_today").set_value(None).run()
        self.assertEqual(len(at.exception), 0)
        self.assertEqual(at.session_state["oms_dash_active_today"], date.today())
        self.assertEqual(
            at.session_state["oms_dash_df"].attrs.get("today_date"),
            date.today().strftime("%Y-%m-%d"),
        )

    def test_interactive_filters_search_reset_and_clear(self):
        ref_df = _build_reference_oms_df()
        at = self._new_app()
        self._upload_oms_and_set_today(
            at,
            "Nov1-25-OMS.csv",
            _df_to_csv_bytes(ref_df),
            "text/csv",
            date(2026, 11, 25),
        )

        on_hold_opt = next(o for o in at.pills(key="oms_dash_filter_pill").options if o.startswith("On Hold"))
        at.pills(key="oms_dash_filter_pill").select(on_hold_opt).run()
        self.assertEqual(len(at.exception), 0)
        tracker_df = next(d.value for d in at.dataframe if "Ticket ID" in d.value.columns and "Status" in d.value.columns)
        self.assertEqual(len(tracker_df), 1)
        self.assertEqual(tracker_df.iloc[0]["Ticket ID"], "EX-2026-46397")

        at.button(key="oms_dash_filter_reset").click().run()
        self.assertEqual(len(at.exception), 0)
        tracker_df = next(d.value for d in at.dataframe if "Ticket ID" in d.value.columns and "Status" in d.value.columns)
        self.assertEqual(len(tracker_df), 9)

        at.text_input(key="oms_dash_filter_search").input("EX-2026-32559").run()
        self.assertEqual(len(at.exception), 0)
        tracker_df = next(d.value for d in at.dataframe if "Ticket ID" in d.value.columns and "Status" in d.value.columns)
        self.assertEqual(len(tracker_df), 1)
        self.assertEqual(tracker_df.iloc[0]["Ticket ID"], "EX-2026-32559")

        at.button(key="oms_dash_filter_reset").click().run()
        self.assertEqual(len(at.exception), 0)
        tracker_df = next(d.value for d in at.dataframe if "Ticket ID" in d.value.columns and "Status" in d.value.columns)
        self.assertEqual(len(tracker_df), 9)

        at.multiselect(key="oms_dash_filter_zones").select("South").run()
        self.assertEqual(len(at.exception), 0)
        tracker_df = next(d.value for d in at.dataframe if "Ticket ID" in d.value.columns and "Status" in d.value.columns)
        self.assertEqual(len(tracker_df), 2)

        at.sidebar.button(key="oms_dash_clear_btn").click().run()
        self.assertEqual(len(at.exception), 0)
        self.assertIsNone(at.session_state["oms_dash_df"])

    def test_subset_24col_oms_dump_resilience(self):
        if not SAMPLE_OMS_PATH.exists():
            self.skipTest(f"Sample OMS dump not found: {SAMPLE_OMS_PATH}")

        at = self._new_app()
        raw_bytes = SAMPLE_OMS_PATH.read_bytes()
        at.sidebar.file_uploader(key="oms_dash_uploader").upload(
            "OMS DUMP.xlsx",
            raw_bytes,
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        ).run()
        self.assertEqual(len(at.exception), 0, f"Exceptions on 24-col OMS DUMP: {at.exception}")

        conv_df = at.session_state["oms_dash_df"]
        self.assertIsNotNone(conv_df)
        self.assertEqual(len(conv_df), 2605)
        self.assertEqual(len(conv_df.columns), 24)
        self.assertGreaterEqual(len(at.get("vega_lite_chart")), 4)
        tracker_df = next(d.value for d in at.dataframe if "Ticket ID" in d.value.columns and "Status" in d.value.columns)
        self.assertEqual(len(tracker_df), 2605)

    def test_view_isolation_and_compare_regression(self):
        if not (SAMPLE_FIR_PATH.exists() and SAMPLE_OMS_PATH.exists()):
            self.skipTest("Sample DUMP files missing")

        ref_df = _build_reference_oms_df()
        at = self._new_app()
        self._upload_oms_and_set_today(
            at,
            "Nov1-25-OMS.csv",
            _df_to_csv_bytes(ref_df),
            "text/csv",
            date(2026, 12, 5),
        )

        at.sidebar.radio(key="nav_view").set_value("OMS-FIR Compare").run()
        self.assertEqual(len(at.exception), 0)
        self.assertIsNone(at.session_state["fir_df"])
        self.assertIsNone(at.session_state["oms_df"])
        self.assertIsNone(at.session_state["reconciliation_result"])

        sample_btn = next(b for b in at.sidebar.button if b.label == "Sample Files")
        sample_btn.click().run()
        self.assertEqual(len(at.exception), 0)
        self.assertIsNotNone(at.session_state["fir_df"])
        self.assertIsNotNone(at.session_state["oms_df"])

        compare_btn = next(b for b in at.sidebar.button if b.label == "Compare Both Files")
        compare_btn.click().run()
        self.assertEqual(len(at.exception), 0)
        self.assertIsNotNone(at.session_state["reconciliation_result"])
        self.assertEqual(len(at.tabs), 6)
        self.assertEqual(len(at.get("vega_lite_chart")), 4)
        self.assertEqual(len(at.dataframe), 5)
        self.assertEqual(len(at.download_button), 6)

        at.sidebar.radio(key="nav_view").set_value("OMS Dashboard").run()
        self.assertEqual(len(at.exception), 0)
        self.assertIsNotNone(at.session_state["oms_dash_df"])
        self.assertEqual(len(at.session_state["oms_dash_df"]), 9)
        self.assertEqual(at.session_state["oms_dash_active_today"], date(2026, 12, 5))
        self.assertEqual(at.sidebar.date_input(key="oms_dash_today").value, date(2026, 12, 5))

    def test_views_import_safety_and_app_reexports(self):
        code = (
            "import warnings; "
            "import app, views, views.styles, views.compare_view, views.oms_dashboard; "
            "required = ['SAMPLE_FIR_PATH', 'SAMPLE_OMS_PATH', 'init_session_state', "
            "'render_landing_view', 'render_audit_view', 'calc_table_height', "
            "'CATEGORY_GAP', 'CATEGORY_PHANTOM', 'CATEGORY_ACTIVE', 'CATEGORY_HEALTHY', "
            "'ReconciliationResult', 'ReconciliationSummary', 'load_dataset', "
            "'reconcile_incident_records', 'create_filtered_result', "
            "'generate_reconciliation_excel', 'export_reconciliation_csv']; "
            "missing = [k for k in required if not hasattr(app, k)]; "
            "assert not missing, f'Missing re-exports: {missing}'; "
            "print('IMPORT_OK')"
        )
        proc = subprocess.run(
            [sys.executable, "-c", code],
            cwd=str(PROJECT_ROOT),
            capture_output=True,
            text=True,
            check=True,
        )
        self.assertIn("IMPORT_OK", proc.stdout)
        self.assertNotIn("ScriptRunContext", proc.stderr)
        self.assertNotIn("ScriptRunContext", proc.stdout)

    def test_empty_and_invalid_uploads_graceful_handling(self):
        at = self._new_app()

        empty_csv = (",".join(NOV1_25_OMS_COLUMNS) + "\n").encode("utf-8")
        at.sidebar.file_uploader(key="oms_dash_uploader").upload(
            "empty_oms.csv", empty_csv, "text/csv"
        ).run()
        self.assertEqual(len(at.exception), 0)
        self.assertIsNotNone(at.session_state["oms_dash_df"])
        self.assertEqual(len(at.session_state["oms_dash_df"]), 0)
        self.assertGreaterEqual(len(at.info), 4)

        at.sidebar.file_uploader(key="oms_dash_uploader").upload(
            "corrupted.xlsx", b"not-a-valid-xlsx-file",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        ).run()
        self.assertEqual(len(at.exception), 0)
        self.assertGreaterEqual(len(at.sidebar.error), 1)
        self.assertEqual(len(at.session_state["oms_dash_df"]), 0)

    def test_compare_sample_after_upload_not_overwritten(self):
        if not (SAMPLE_FIR_PATH.exists() and SAMPLE_OMS_PATH.exists()):
            self.skipTest("Sample DUMP files missing")

        at = self._new_app()
        at.sidebar.radio(key="nav_view").set_value("OMS-FIR Compare").run()
        self.assertEqual(len(at.exception), 0)

        mini_fir = pd.DataFrame(
            [{"Station Name": "S1", "OCPP ID": "CHG-1", "Status": "Faulted", "Issue_Age": "1 day"}]
        )
        at.sidebar.file_uploader(key="fir_uploader").upload(
            "mini_fir.csv", _df_to_csv_bytes(mini_fir), "text/csv"
        ).run()
        self.assertEqual(len(at.exception), 0)
        self.assertEqual(len(at.session_state["fir_df"]), 1)

        sample_btn = next(b for b in at.sidebar.button if b.label == "Sample Files")
        sample_btn.click().run()
        self.assertEqual(len(at.exception), 0)
        self.assertEqual(at.session_state["dataset_source"], "Sample Demo Files")
        self.assertEqual(len(at.session_state["fir_df"]), 54)
        self.assertEqual(len(at.session_state["oms_df"]), 2605)

    def test_compare_search_regex_metacharacters_and_html_escaping(self):
        if not (SAMPLE_FIR_PATH.exists() and SAMPLE_OMS_PATH.exists()):
            self.skipTest("Sample DUMP files missing")

        at = self._new_app()
        at.sidebar.radio(key="nav_view").set_value("OMS-FIR Compare").run()
        self.assertEqual(len(at.exception), 0)

        next(b for b in at.sidebar.button if b.label == "Sample Files").click().run()
        self.assertEqual(len(at.exception), 0)
        next(b for b in at.sidebar.button if b.label == "Compare Both Files").click().run()
        self.assertEqual(len(at.exception), 0)

        with warnings.catch_warnings(record=True) as caught_warnings:
            warnings.simplefilter("always")
            for query in (".*+?()[]", "[", "<img src=x onerror=alert(1)>"):
                at.text_input(key="dyn_search_input").input(query).run()
                self.assertEqual(len(at.exception), 0, f"Exception on query {query!r}: {at.exception}")

        regex_warnings = [
            w for w in caught_warnings
            if issubclass(w.category, UserWarning) and "regular expression" in str(w.message).lower()
        ]
        self.assertEqual(len(regex_warnings), 0, f"Unexpected regex warnings: {[str(w.message) for w in regex_warnings]}")
        self.assertEqual(len(caught_warnings), 0, f"Unexpected warnings: {[str(w.message) for w in caught_warnings]}")

        md_values = [m.value for m in at.markdown]
        self.assertTrue(
            any("&lt;img src=x onerror=alert(1)&gt;" in m for m in md_values),
            "Expected escaped HTML tag in active filter summary markdown",
        )
        self.assertFalse(
            any("<img src=x onerror=alert(1)>" in m for m in md_values),
            "Unescaped <img> tag found in markdown output",
        )

    def test_oms_dashboard_uploaded_file_with_existing_no_column(self):
        at = self._new_app()
        csv_with_no = b"No.,Ticket ID,Status\n10,T-1,Open\n20,T-2,Closed\n"
        at.sidebar.file_uploader(key="oms_dash_uploader").upload(
            "with_no_col.csv", csv_with_no, "text/csv"
        ).run()
        self.assertEqual(len(at.exception), 0, f"Exception with No. column: {at.exception}")
        self.assertEqual(len(at.dataframe), 2)
        tracker_df = at.dataframe[0].value
        schema_df = at.dataframe[1].value
        self.assertEqual(len(tracker_df), 2)
        self.assertEqual(len(schema_df), 2)
        self.assertEqual(list(schema_df["No."]), [1, 2])
        self.assertIn("Original No.", schema_df.columns)
        self.assertEqual([str(v) for v in schema_df["Original No."]], ["10", "20"])
        self.assertEqual(list(at.session_state["oms_dash_df"].columns), ["No.", "Ticket ID", "Status"])

        csv_fallback_no = b"No.,Note\n5,Alpha\n6,Beta\n"
        at.sidebar.file_uploader(key="oms_dash_uploader").upload(
            "fallback_no_col.csv", csv_fallback_no, "text/csv"
        ).run()
        self.assertEqual(len(at.exception), 0, f"Exception on fallback tracker with No. column: {at.exception}")
        self.assertEqual(len(at.dataframe), 2)
        self.assertEqual(list(at.dataframe[0].value["No."]), [1, 2])
        self.assertEqual([str(v) for v in at.dataframe[0].value["Original No."]], ["5", "6"])

    def test_oms_dashboard_whitespace_only_chart_labels(self):
        at = self._new_app()
        ws_df = pd.DataFrame(
            [
                {
                    "Ticket ID": "T-WS-1",
                    "Status": "  ",
                    "Zone": "   ",
                    "Priority": "\t",
                    "SLA Breached": "  ",
                    "Issue Category": "   ",
                    "Station": "  ",
                },
                {
                    "Ticket ID": "T-WS-2",
                    "Status": "\t ",
                    "Zone": " ",
                    "Priority": "  ",
                    "SLA Breached": "\t",
                    "Issue Category": " ",
                    "Station": "\t",
                },
            ]
        )
        at.sidebar.file_uploader(key="oms_dash_uploader").upload(
            "whitespace_labels.csv", _df_to_csv_bytes(ws_df), "text/csv"
        ).run()
        self.assertEqual(len(at.exception), 0, f"Exception on whitespace chart labels: {at.exception}")

        charts = at.get("vega_lite_chart")[:4]
        self.assertEqual(len(charts), 4)

        expected_fields = [
            {"Status_Clean": "Unspecified", "Zone_Clean": "Unspecified"},
            {"Priority_Clean": "Unspecified", "SLA_Clean": "Unknown"},
            {"Category_Clean": "Unspecified"},
            {"Station_Clean": "Unspecified"},
        ]
        for chart_el, field_map in zip(charts, expected_fields):
            self.assertGreater(len(chart_el.proto.datasets), 0)
            chart_df = pa_ipc.open_stream(chart_el.proto.datasets[0].data.data).read_all().to_pandas()
            self.assertGreater(len(chart_df), 0)
            for field, expected_label in field_map.items():
                self.assertIn(field, chart_df.columns)
                vals = chart_df[field].tolist()
                self.assertEqual(vals, [expected_label] * len(vals), f"Unexpected labels for {field}: {vals}")


    def test_mpr_top10_categories_and_top5_subtypes(self):
        ref_df = _build_reference_oms_df()
        at = self._new_app()
        self._upload_oms_and_set_today(
            at,
            "Nov1-25-OMS.xlsx",
            _df_to_xlsx_bytes(ref_df),
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            date(2026, 11, 25),
        )
        self.assertEqual(len(at.exception), 0)
        top_cat_df = next((d.value for d in at.dataframe if "Issue Category" in d.value.columns and "Total Tickets" in d.value.columns), None)
        self.assertIsNotNone(top_cat_df, "Expected Top 10 Issue Categories table")
        self.assertIn("Total Tickets", top_cat_df.columns)
        self.assertIn("Share %", top_cat_df.columns)
        self.assertGreater(len(top_cat_df), 0)

    def test_mpr_station_wise_repetitive_issues(self):
        ref_df = _build_reference_oms_df()
        at = self._new_app()
        self._upload_oms_and_set_today(
            at,
            "Nov1-25-OMS.xlsx",
            _df_to_xlsx_bytes(ref_df),
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            date(2026, 11, 25),
        )
        self.assertEqual(len(at.exception), 0)
        top_st_df = next((d.value for d in at.dataframe if "Station" in d.value.columns and "Top Failure Mode" in d.value.columns), None)
        self.assertIsNotNone(top_st_df, "Expected Top 10 Station Hotspots table")
        self.assertIn("Unique Chargers", top_st_df.columns)

    def test_mpr_engineer_performance_scorecard(self):
        ref_df = _build_reference_oms_df()
        at = self._new_app()
        self._upload_oms_and_set_today(
            at,
            "Nov1-25-OMS.xlsx",
            _df_to_xlsx_bytes(ref_df),
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            date(2026, 11, 25),
        )
        self.assertEqual(len(at.exception), 0)
        scorecard_df = next((d.value for d in at.dataframe if "Engineer" in d.value.columns and "Resolution Rate %" in d.value.columns), None)
        self.assertIsNotNone(scorecard_df, "Expected Field Engineer Scorecard table")
        self.assertIn("SLA Adherence %", scorecard_df.columns)
        self.assertIn("Total Assigned", scorecard_df.columns)

if __name__ == "__main__":
    unittest.main()
