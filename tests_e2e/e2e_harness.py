"""Opaque-box Streamlit AppTest harness and public-contract helpers."""

from __future__ import annotations

from contextlib import contextmanager
from datetime import date
from html.parser import HTMLParser
import io
import json
from pathlib import Path
from typing import Any, Callable, Dict, Iterable, List, Optional, Sequence
from unittest.mock import patch

import openpyxl
import pandas as pd
from streamlit.runtime.media_file_manager import MediaFileManager
from streamlit.testing.v1 import AppTest

PROJECT_ROOT = Path(__file__).resolve().parent.parent
APP_PATH = PROJECT_ROOT / "app.py"
DEFAULT_TIMEOUT = 30

VIEW_OMS_DASHBOARD = "OMS Dashboard"
VIEW_OMS_FIR_COMPARE = "OMS-FIR Compare"


class _HTMLTextExtractor(HTMLParser):
    """Strip HTML tags and ignore <style>/<script> contents."""

    def __init__(self) -> None:
        super().__init__()
        self._skip_depth = 0
        self._chunks: List[str] = []

    def handle_starttag(self, tag: str, attrs: list) -> None:
        if tag.lower() in ("style", "script"):
            self._skip_depth += 1

    def handle_endtag(self, tag: str) -> None:
        if tag.lower() in ("style", "script") and self._skip_depth > 0:
            self._skip_depth -= 1
        elif tag.lower() in ("div", "p", "br", "li", "tr", "h1", "h2", "h3", "h4"):
            self._chunks.append("\n")

    def handle_data(self, data: str) -> None:
        if self._skip_depth == 0:
            self._chunks.append(data)

    def get_text(self) -> str:
        raw = "".join(self._chunks)
        lines = [" ".join(line.split()) for line in raw.splitlines()]
        return "\n".join(line for line in lines if line)


def html_to_text(html_str: str) -> str:
    parser = _HTMLTextExtractor()
    parser.feed(str(html_str or ""))
    parser.close()
    return parser.get_text()


_VOID_HTML_TAGS = frozenset(
    {"area", "base", "br", "col", "embed", "hr", "img", "input", "link", "meta", "param", "source", "track", "wbr"}
)
_BLOCK_HTML_TAGS = frozenset({"div", "p", "br", "li", "tr", "h1", "h2", "h3", "h4", "h5", "h6", "section", "article"})


class _KPICardExtractor(HTMLParser):
    """Extract each .kpi-card block (or .kpi-title group) as a single normalized line."""

    def __init__(self, *, group_by_title: bool = False) -> None:
        super().__init__()
        self._group_by_title = group_by_title
        self._skip_depth = 0
        self._stack: List[tuple[str, set[str]]] = []
        self._card_depth: Optional[int] = None
        self._segments: List[str] = []
        self._cur_buf: List[str] = []
        self.cards: List[str] = []

    def _flush_segment(self) -> None:
        text = " ".join("".join(self._cur_buf).split())
        self._cur_buf.clear()
        if text:
            self._segments.append(text)

    def _finish_card(self) -> None:
        self._flush_segment()
        if self._segments:
            if len(self._segments) == 1:
                self.cards.append(self._segments[0])
            else:
                self.cards.append(f"{self._segments[0]}: {' '.join(self._segments[1:])}")
        self._segments.clear()
        self._card_depth = None

    def handle_starttag(self, tag: str, attrs: list) -> None:
        t = tag.lower()
        if t in ("style", "script"):
            self._skip_depth += 1
            return
        if self._skip_depth > 0:
            return
        if t in _VOID_HTML_TAGS:
            if self._card_depth is not None and t == "br":
                self._flush_segment()
            return
        attr_dict = dict(attrs)
        classes = set(str(attr_dict.get("class") or "").split())
        self._stack.append((t, classes))
        if self._group_by_title:
            if "kpi-title" in classes:
                if self._card_depth is not None:
                    self._finish_card()
                self._card_depth = 1
                self._segments = []
                self._cur_buf = []
            elif self._card_depth is not None and (
                t in _BLOCK_HTML_TAGS or any(c.startswith("kpi-") for c in classes)
            ):
                self._flush_segment()
            return
        if self._card_depth is None and "kpi-card" in classes:
            self._card_depth = len(self._stack)
            self._segments = []
            self._cur_buf = []
        elif self._card_depth is not None and (
            t in _BLOCK_HTML_TAGS or any(c.startswith("kpi-") for c in classes)
        ):
            self._flush_segment()

    def handle_endtag(self, tag: str) -> None:
        t = tag.lower()
        if t in ("style", "script") and self._skip_depth > 0:
            self._skip_depth -= 1
            return
        if self._skip_depth > 0 or t in _VOID_HTML_TAGS:
            return
        idx = None
        for i in range(len(self._stack) - 1, -1, -1):
            if self._stack[i][0] == t:
                idx = i
                break
        if idx is None:
            return
        closing_depth = idx + 1
        closed_tag, closed_classes = self._stack[idx]
        del self._stack[idx:]
        if self._card_depth is not None:
            if not self._group_by_title and closing_depth <= self._card_depth:
                self._finish_card()
            elif closed_tag in _BLOCK_HTML_TAGS or any(c.startswith("kpi-") for c in closed_classes):
                self._flush_segment()

    def handle_data(self, data: str) -> None:
        if self._skip_depth == 0 and self._card_depth is not None:
            self._cur_buf.append(data)

    def close(self) -> None:
        super().close()
        if self._card_depth is not None:
            self._finish_card()


def extract_kpi_cards_from_html(html_str: str) -> List[str]:
    raw = str(html_str or "")
    extractor = _KPICardExtractor(group_by_title=False)
    extractor.feed(raw)
    extractor.close()
    if extractor.cards:
        return extractor.cards
    if "kpi-title" in raw:
        title_extractor = _KPICardExtractor(group_by_title=True)
        title_extractor.feed(raw)
        title_extractor.close()
        if title_extractor.cards:
            return title_extractor.cards
    fallback = " ".join(html_to_text(raw).splitlines())
    return [fallback] if fallback else []


class AppSession:
    """Wrapper around Streamlit AppTest with media download interception."""

    def __init__(self, app_path: Path = APP_PATH, timeout: int = DEFAULT_TIMEOUT) -> None:
        self.app_path = Path(app_path)
        self.timeout = timeout
        self.at: AppTest = AppTest.from_file(str(self.app_path), default_timeout=timeout)
        self.media_map: Dict[str, Dict[str, Any]] = {}

    @contextmanager
    def capture_downloads(self):
        orig_add = MediaFileManager.add
        session_media = self.media_map

        def _wrapped_add(
            mgr_self,
            path_or_data,
            mimetype: str,
            coordinates: str,
            file_name: Optional[str] = None,
            is_for_static_download: bool = False,
        ):
            url = orig_add(
                mgr_self,
                path_or_data,
                mimetype,
                coordinates,
                file_name=file_name,
                is_for_static_download=is_for_static_download,
            )
            if isinstance(path_or_data, (bytes, bytearray)):
                raw_bytes = bytes(path_or_data)
            elif isinstance(path_or_data, str):
                p = Path(path_or_data)
                if p.exists() and p.is_file():
                    raw_bytes = p.read_bytes()
                else:
                    raw_bytes = path_or_data.encode("utf-8")
            else:
                raw_bytes = bytes(path_or_data)

            session_media[url] = {
                "url": url,
                "file_name": file_name or "",
                "mimetype": mimetype or "",
                "bytes": raw_bytes,
            }
            return url

        with patch.object(MediaFileManager, "add", autospec=True, side_effect=_wrapped_add):
            yield

    def start(self) -> "AppSession":
        with self.capture_downloads():
            self.at.run(timeout=self.timeout)
        return self

    def rerun(self) -> "AppSession":
        with self.capture_downloads():
            self.at.run(timeout=self.timeout)
        return self

    def run_widget_action(self, action_fn: Callable[[], Any]) -> "AppSession":
        with self.capture_downloads():
            res = action_fn()
            if hasattr(res, "run"):
                res.run(timeout=self.timeout)
        return self

    def set_session_state(self, key: str, value: Any) -> "AppSession":
        self.at.session_state[key] = value
        return self

    def assert_no_exceptions(self) -> None:
        if len(self.at.exception) > 0:
            msgs = [f"{exc.message}" for exc in self.at.exception]
            raise AssertionError("Streamlit app raised unexpected exception(s): " + " | ".join(msgs))

    def get_nav_radio(self):
        candidates = list(self.at.sidebar.radio) + list(self.at.radio)
        for radio in candidates:
            opts = [str(o) for o in getattr(radio, "options", [])]
            if VIEW_OMS_DASHBOARD in opts and VIEW_OMS_FIR_COMPARE in opts:
                return radio
        raise AssertionError(
            f"Sidebar navigation radio with options [{VIEW_OMS_DASHBOARD!r}, {VIEW_OMS_FIR_COMPARE!r}] not found."
        )

    def switch_view(self, view_name: str, check_exceptions: bool = True) -> "AppSession":
        radio = self.get_nav_radio()
        opts = [str(o) for o in radio.options]
        if view_name not in opts:
            raise AssertionError(f"View {view_name!r} not in navigation options: {opts}")
        self.run_widget_action(lambda: radio.set_value(view_name))
        if check_exceptions:
            self.assert_no_exceptions()
        return self

    def find_dashboard_uploader(self):
        """Find the OMS Dashboard file uploader without matching OMS-FIR Compare uploaders."""
        compare_keys = {"fir_uploader", "oms_uploader"}
        uploaders = list(self.at.sidebar.file_uploader) + list(self.at.file_uploader)
        seen_keys = set()
        non_compare = []
        for up in uploaders:
            k = getattr(up, "key", None)
            if k in seen_keys:
                continue
            seen_keys.add(k)
            if k in compare_keys:
                continue
            lbl = str(getattr(up, "label", "")).lower()
            if "field report" in lbl or "helpdesk tickets (oms)" in lbl:
                continue
            non_compare.append(up)
        for up in non_compare:
            k = str(getattr(up, "key", "")).lower()
            lbl = str(getattr(up, "label", "")).lower()
            if "oms" in k or "oms" in lbl or "dash" in k or "dump" in lbl:
                return up
        if len(non_compare) == 1:
            return non_compare[0]
        raise AssertionError(
            "OMS Dashboard file_uploader not found (distinct from Compare fir_uploader/oms_uploader)."
        )

    def find_compare_uploaders(self) -> tuple[Any, Any]:
        uploaders = list(self.at.sidebar.file_uploader) + list(self.at.file_uploader)
        fir_up = None
        oms_up = None
        for up in uploaders:
            k = str(getattr(up, "key", ""))
            lbl = str(getattr(up, "label", "")).lower()
            if k == "fir_uploader" or "field report" in lbl:
                fir_up = up
            elif k == "oms_uploader" or "helpdesk tickets" in lbl:
                oms_up = up
        if fir_up is None or oms_up is None:
            raise AssertionError("OMS-FIR Compare uploaders (fir_uploader, oms_uploader) not found.")
        return fir_up, oms_up

    def upload_to(
        self,
        uploader,
        filename: str,
        content: bytes,
        mime_type: str = "application/octet-stream",
    ) -> "AppSession":
        return self.run_widget_action(lambda: uploader.upload(filename, content, mime_type))

    def clear_uploader(self, uploader) -> "AppSession":
        return self.run_widget_action(lambda: uploader.clear())

    def find_today_date_input(self):
        seen = set()
        inputs = []
        for di in list(self.at.sidebar.date_input) + list(self.at.date_input):
            uid = getattr(getattr(di, "proto", None), "id", None) or getattr(di, "key", None) or id(di)
            if uid not in seen:
                seen.add(uid)
                inputs.append(di)

        def _is_single_date(di) -> bool:
            return not isinstance(getattr(di, "value", None), (list, tuple))

        def _matches_today(di) -> bool:
            lbl = str(getattr(di, "label", "")).lower()
            k = str(getattr(di, "key", "")).lower()
            return "today" in lbl or "reference" in lbl or "today" in k

        single_inputs = [di for di in inputs if _is_single_date(di)]
        for di in single_inputs:
            if _matches_today(di):
                return di
        for di in inputs:
            if _matches_today(di):
                return di
        if len(single_inputs) == 1:
            return single_inputs[0]
        if len(inputs) == 1:
            return inputs[0]
        raise AssertionError("OMS Dashboard 'today' st.date_input control not found.")

    def set_date(self, date_input, value: date) -> "AppSession":
        return self.run_widget_action(lambda: date_input.set_value(value))

    def find_button(self, label_contains: str):
        needle = label_contains.lower()
        for btn in list(self.at.sidebar.button) + list(self.at.button):
            if needle in str(getattr(btn, "label", "")).lower():
                return btn
        raise AssertionError(f"Button containing {label_contains!r} not found.")

    def click_button(self, label_contains: str) -> "AppSession":
        btn = self.find_button(label_contains)
        return self.run_widget_action(lambda: btn.click())

    def find_search_input(self):
        inputs = list(self.at.text_input) + list(self.at.sidebar.text_input)
        for ti in inputs:
            k = str(getattr(ti, "key", "")).lower()
            lbl = str(getattr(ti, "label", "")).lower()
            ph = str(getattr(ti, "placeholder", "")).lower()
            if any(tok in (k + " " + lbl + " " + ph) for tok in ("search", "filter", "ticket", "query", "find")):
                return ti
        if inputs:
            return inputs[0]
        raise AssertionError("Search st.text_input not found.")

    def set_text(self, text_input, value: str) -> "AppSession":
        return self.run_widget_action(lambda: text_input.set_value(value))

    def find_multiselects(self) -> List[Any]:
        seen = set()
        result = []
        for ms in list(self.at.multiselect) + list(self.at.sidebar.multiselect):
            uid = getattr(getattr(ms, "proto", None), "id", None) or getattr(ms, "key", None) or id(ms)
            if uid not in seen:
                seen.add(uid)
                result.append(ms)
        return result

    def find_multiselect(self, label_or_option_contains: str):
        needle = label_or_option_contains.lower()
        for ms in self.find_multiselects():
            lbl = str(getattr(ms, "label", "")).lower()
            k = str(getattr(ms, "key", "")).lower()
            opts = [str(o).lower() for o in getattr(ms, "options", [])]
            if needle in lbl or needle in k or any(needle in o for o in opts):
                return ms
        raise AssertionError(f"Multiselect matching {label_or_option_contains!r} not found.")

    def set_multiselect(self, ms, values: Sequence[str]) -> "AppSession":
        return self.run_widget_action(lambda: ms.set_value(list(values)))

    def find_selectboxes(self) -> List[Any]:
        seen = set()
        result = []
        for sb in list(self.at.selectbox) + list(self.at.sidebar.selectbox):
            uid = getattr(getattr(sb, "proto", None), "id", None) or getattr(sb, "key", None) or id(sb)
            if uid not in seen:
                seen.add(uid)
                result.append(sb)
        return result

    def set_selectbox(self, sb, value: Any) -> "AppSession":
        return self.run_widget_action(lambda: sb.set_value(value))

    def find_pills(self) -> List[Any]:
        seen = set()
        result = []
        for p in list(self.at.pills) + list(self.at.sidebar.pills):
            uid = getattr(getattr(p, "proto", None), "id", None) or getattr(p, "key", None) or id(p)
            if uid not in seen:
                seen.add(uid)
                result.append(p)
        return result

    @staticmethod
    def _is_multi_pills(pill) -> bool:
        if isinstance(getattr(pill, "value", None), list):
            return True
        if getattr(pill, "selection_mode", None) == "multi":
            return True
        return getattr(getattr(pill, "proto", None), "click_mode", None) == 1

    def select_pill_containing(self, substring: str) -> "AppSession":
        needle = substring.lower()
        pills = self.find_pills()
        for exact_only in (True, False):
            for p in pills:
                for opt in getattr(p, "options", []):
                    opt_s = str(opt)
                    matched = (opt_s.strip().lower() == needle) if exact_only else (needle in opt_s.lower())
                    if matched:
                        if self._is_multi_pills(p):
                            current = list(getattr(p, "value", []) or [])
                            target = current if opt in current else current + [opt]
                            return self.run_widget_action(lambda w=p, v=target: w.set_value(v))
                        return self.run_widget_action(lambda w=p, v=opt: w.set_value(v))
        raise AssertionError(f"Pill option containing {substring!r} not found.")

    def select_dashboard_filter(self, option_value: str) -> "AppSession":
        """Apply a dashboard filter across st.pills, st.multiselect, or st.selectbox."""
        needle = option_value.lower()
        pills = self.find_pills()
        multiselects = self.find_multiselects()
        selectboxes = self.find_selectboxes()

        for exact_only in (True, False):
            for p in pills:
                for opt in getattr(p, "options", []):
                    opt_s = str(opt)
                    matched = (opt_s.strip().lower() == needle) if exact_only else (needle in opt_s.lower())
                    if matched:
                        if self._is_multi_pills(p):
                            current = list(getattr(p, "value", []) or [])
                            target = current if opt in current else current + [opt]
                            self.run_widget_action(lambda w=p, v=target: w.set_value(v))
                        else:
                            self.run_widget_action(lambda w=p, v=opt: w.set_value(v))
                        self.assert_no_exceptions()
                        return self

            for ms in multiselects:
                for opt in getattr(ms, "options", []):
                    opt_s = str(opt)
                    matched = (opt_s.strip().lower() == needle) if exact_only else (needle in opt_s.lower())
                    if matched:
                        current = list(getattr(ms, "value", []) or [])
                        target = current if opt in current else current + [opt]
                        self.set_multiselect(ms, target)
                        self.assert_no_exceptions()
                        return self

            for sb in selectboxes:
                for opt in getattr(sb, "options", []):
                    opt_s = str(opt)
                    matched = (opt_s.strip().lower() == needle) if exact_only else (needle in opt_s.lower())
                    if matched:
                        self.set_selectbox(sb, opt)
                        self.assert_no_exceptions()
                        return self

        raise AssertionError(f"No filter control found containing option {option_value!r}.")

    def clear_dashboard_filters(self) -> "AppSession":
        """Reset active search box, reset button, pills, multiselects, and selectboxes."""
        seen_btns = set()
        buttons = []
        for btn in list(self.at.button) + list(self.at.sidebar.button):
            uid = getattr(getattr(btn, "proto", None), "id", None) or getattr(btn, "key", None) or id(btn)
            if uid not in seen_btns:
                seen_btns.add(uid)
                buttons.append(btn)
        for btn in buttons:
            lbl = str(getattr(btn, "label", "")).lower()
            if "reset" in lbl and "clear all" not in lbl:
                self.run_widget_action(lambda b=btn: b.click())
                self.assert_no_exceptions()

        try:
            ti = self.find_search_input()
        except AssertionError:
            ti = None
        if ti is not None and getattr(ti, "value", "") != "":
            self.set_text(ti, "")

        for p in self.find_pills():
            if self._is_multi_pills(p):
                if getattr(p, "value", None):
                    self.run_widget_action(lambda w=p: w.set_value([]))
            else:
                opts = list(getattr(p, "options", []))
                all_opts = [o for o in opts if str(o).strip().lower() in ("all", "all tickets")]
                if all_opts:
                    if getattr(p, "value", None) != all_opts[0]:
                        self.run_widget_action(lambda w=p, v=all_opts[0]: w.set_value(v))
                elif getattr(p, "value", None) is not None:
                    self.run_widget_action(lambda w=p: w.set_value([]))

        for ms in self.find_multiselects():
            if getattr(ms, "value", None):
                self.set_multiselect(ms, [])

        for sb in self.find_selectboxes():
            opts = list(getattr(sb, "options", []))
            if opts:
                all_opts = [o for o in opts if str(o).strip().lower() in ("all", "all tickets")]
                target = all_opts[0] if all_opts else opts[0]
                if getattr(sb, "value", None) != target:
                    self.set_selectbox(sb, target)

        self.assert_no_exceptions()
        return self

    def get_downloads(self) -> List[Dict[str, Any]]:
        downloads = []
        for dl in list(self.at.download_button) + list(self.at.sidebar.download_button):
            url = getattr(dl, "url", "")
            meta = self.media_map.get(url, {})
            downloads.append(
                {
                    "label": str(getattr(dl, "label", "")),
                    "url": url,
                    "file_name": meta.get("file_name", ""),
                    "mimetype": meta.get("mimetype", ""),
                    "bytes": meta.get("bytes", b""),
                }
            )
        return downloads

    def find_download(
        self,
        *,
        suffix: Optional[str] = None,
        label_contains: Optional[str] = None,
        filename_contains: Optional[str] = None,
    ) -> Dict[str, Any]:
        matches = []
        for item in self.get_downloads():
            fname = item["file_name"].lower()
            lbl = item["label"].lower()
            if suffix and not (fname.endswith(suffix.lower()) or suffix.lower().lstrip(".") in lbl):
                continue
            if label_contains and label_contains.lower() not in lbl:
                continue
            if filename_contains and filename_contains.lower() not in fname:
                continue
            matches.append(item)

        if not matches:
            raise AssertionError(
                f"Download button not found (suffix={suffix!r}, label={label_contains!r}, file={filename_contains!r}). "
                f"Available: {[(d['label'], d['file_name']) for d in self.get_downloads()]}"
            )

        if len(matches) > 1 and label_contains is None and filename_contains is None:
            non_filtered = [
                m
                for m in matches
                if "filtered" not in m["label"].lower() and "filtered" not in m["file_name"].lower()
            ]
            pool = non_filtered or matches
            for tok in ("converted", "date_only", "full", "all"):
                for m in pool:
                    if tok in m["label"].lower() or tok in m["file_name"].lower():
                        return m
            return pool[0]

        return matches[0]

    def get_dataframes(self) -> List[pd.DataFrame]:
        return [df.value for df in self.at.dataframe]

    def get_charts(self) -> List[Dict[str, Any]]:
        charts = []
        for el in self.at.get("vega_lite_chart"):
            spec_str = getattr(el.proto, "spec", "{}")
            try:
                spec = json.loads(spec_str)
            except json.JSONDecodeError:
                spec = {}
            datasets = {}
            for ds in getattr(el.proto, "datasets", []):
                ds_name = getattr(ds, "name", "")
                ds_data = getattr(ds, "data", None)
                datasets[ds_name] = ds_data
            charts.append({"spec": spec, "datasets": datasets, "raw": el})
        return charts

    def get_chart_signatures(self) -> List[tuple[Any, Dict[str, bytes]]]:
        sigs = []
        for c in self.get_charts():
            ds_bytes: Dict[str, bytes] = {}
            for name, ds_data in c["datasets"].items():
                if hasattr(ds_data, "data"):
                    ds_bytes[name] = bytes(ds_data.data)
                else:
                    ds_bytes[name] = repr(ds_data).encode("utf-8")
            sigs.append((c["spec"], ds_bytes))
        return sigs

    def extract_kpi_cards(self) -> List[str]:
        cards: List[str] = []
        for m in self.at.metric:
            card = " ".join(f"{m.label}: {m.value} {getattr(m, 'delta', '') or ''}".split())
            if card:
                cards.append(card)
        for md in self.at.markdown:
            val = str(md.value)
            if "kpi-card" in val or "kpi-value" in val or "kpi-title" in val:
                cards.extend(extract_kpi_cards_from_html(val))
        return cards

    def extract_kpi_text(self) -> str:
        return "\n".join(self.extract_kpi_cards())

    def extract_all_text(self) -> str:
        parts: List[str] = []
        for col in (
            self.at.title,
            self.at.header,
            self.at.subheader,
            self.at.caption,
            self.at.text,
            self.at.info,
            self.at.success,
            self.at.warning,
            self.at.error,
        ):
            for el in col:
                parts.append(str(getattr(el, "value", "")))
        for m in self.at.metric:
            parts.append(f"{m.label} {m.value}")
        for md in self.at.markdown:
            txt = html_to_text(str(md.value))
            if txt:
                parts.append(txt)
        return "\n".join(parts)

    def get_error_messages(self) -> List[str]:
        return [str(e.value) for e in self.at.error]

    def get_tab_labels(self) -> List[str]:
        return [str(t.label) for t in self.at.tabs]


def parse_csv_bytes(raw_bytes: bytes) -> pd.DataFrame:
    return pd.read_csv(io.BytesIO(raw_bytes), keep_default_na=True)


def load_xlsx_workbook_bytes(raw_bytes: bytes) -> openpyxl.Workbook:
    return openpyxl.load_workbook(io.BytesIO(raw_bytes), data_only=True)


# Lazy public-contract importers (never imported at module collection time)
def get_oms_converter():
    import core.oms_converter as oms_converter

    return oms_converter


def get_oms_exporters():
    from core.exporter import export_oms_date_only_csv, export_oms_date_only_excel

    return export_oms_date_only_csv, export_oms_date_only_excel


def get_core_loader():
    import core.loader as loader

    return loader


def get_core_reconciler():
    import core.reconciler as reconciler

    return reconciler


def assert_reference_converted_df(
    df: pd.DataFrame,
    *,
    expected_hold_map: Optional[Dict[str, Optional[int]]] = None,
) -> None:
    """Assert 42-col schema, 15 YYYY-MM-DD dates, double-space col 28, and all 72 delta cells."""
    from tests_e2e.fixtures import (
        DATE_ONLY_RESOLVED_COL,
        DATETIME_COLS_15,
        DELTA_COLS_8,
        EXPECTED_DELTAS_2026_11_25,
        EXPECTED_STRIPPED_DATES,
        RAW_RESOLVED_COL,
        REFERENCE_42_COLUMNS_DATE_ONLY,
    )

    cols = list(df.columns)
    assert cols == list(REFERENCE_42_COLUMNS_DATE_ONLY), (
        f"42-column schema/order mismatch.\nExpected: {list(REFERENCE_42_COLUMNS_DATE_ONLY)}\nActual:   {cols}"
    )
    assert RAW_RESOLVED_COL not in cols, f"Raw column {RAW_RESOLVED_COL!r} must be renamed."
    assert cols[27] == DATE_ONLY_RESOLVED_COL, f"Col 28 expected {DATE_ONLY_RESOLVED_COL!r}, got {cols[27]!r}"
    assert len(df) == 9, f"Expected 9 reference rows, got {len(df)}"

    by_tid = df.set_index("Ticket ID")
    for tid, expected_dates in EXPECTED_STRIPPED_DATES.items():
        row = by_tid.loc[tid]
        for dcol in DATETIME_COLS_15:
            raw_val = row[dcol]
            actual_str = "" if pd.isna(raw_val) else str(raw_val).strip()
            exp_str = expected_dates[dcol]
            assert actual_str == exp_str, (
                f"[{tid}] Date col {dcol!r}: expected {exp_str!r}, got {actual_str!r}"
            )

        expected_deltas = dict(EXPECTED_DELTAS_2026_11_25[tid])
        if expected_hold_map is not None:
            expected_deltas["Hold total time(Days)"] = expected_hold_map[tid]

        for delta_col in DELTA_COLS_8:
            val = row[delta_col]
            exp_val = expected_deltas[delta_col]
            if exp_val is None:
                is_blank = pd.isna(val) or (isinstance(val, str) and val.strip() == "")
                assert is_blank, f"[{tid}] Delta {delta_col!r}: expected blank, got {val!r}"
            else:
                assert not pd.isna(val), f"[{tid}] Delta {delta_col!r}: expected {exp_val}, got NA"
                assert int(val) == exp_val, f"[{tid}] Delta {delta_col!r}: expected {exp_val}, got {val!r}"
                if isinstance(val, float):
                    assert val.is_integer(), f"[{tid}] Delta {delta_col!r} non-integer float {val!r}"

