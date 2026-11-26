import codecs
import csv
import importlib.util
import io
import re
import struct
from datetime import datetime, timedelta
from html.parser import HTMLParser
from pathlib import Path
from typing import Any, BinaryIO, List, Optional, Set, Tuple, Union
import xml.etree.ElementTree as ET
import pandas as pd

HEADER_KEYWORDS: Set[str] = {
    "ocpp",
    "charger",
    "station",
    "ticket",
    "tenant",
    "priority",
    "status",
    "sr.no",
    "sno",
    "zone",
    "assigned",
    "created",
    "issue",
    "error",
}

_OLE2_MAGIC = b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1"
_OOXML_MAGIC = b"PK\x03\x04"
_BUILTIN_DATE_FORMAT_IDS: Set[int] = {
    14, 15, 16, 17, 18, 19, 20, 21, 22,
    27, 28, 29, 30, 31, 32, 33, 34, 35, 36,
    45, 46, 47, 50, 51, 52, 53, 54, 55, 56, 57, 58,
}


_CSV_FALLBACK_ENCODINGS: Tuple[str, ...] = ("utf-8-sig", "utf-8", "cp1252", "latin-1")


def _safe_seek_zero(stream: Any) -> bool:
    if not hasattr(stream, "seek"):
        return False
    try:
        if hasattr(stream, "seekable") and not stream.seekable():
            return False
        stream.seek(0)
        return True
    except (io.UnsupportedOperation, OSError, ValueError, AttributeError):
        return False


def _ensure_seekable_source(
    source: Union[str, bytes, BinaryIO, Path],
) -> Union[str, bytes, BinaryIO, Path]:
    if hasattr(source, "read") and not _safe_seek_zero(source):
        raw = source.read()
        if isinstance(raw, str):
            return io.BytesIO(raw.encode("utf-8"))
        return io.BytesIO(bytes(raw))
    return source


def _get_bytes_or_buffer(source: Union[str, bytes, BinaryIO, Path]) -> tuple[Any, bool]:
    if isinstance(source, (str, Path)):
        p = Path(source)
        return p, False
    elif isinstance(source, bytes):
        return io.BytesIO(source), True
    elif hasattr(source, "read"):
        if not _safe_seek_zero(source):
            raw = source.read()
            if isinstance(raw, str):
                return io.BytesIO(raw.encode("utf-8")), True
            return io.BytesIO(bytes(raw)), True
        return source, True
    return source, False


def _read_raw_bytes(source: Union[str, bytes, BinaryIO, Path]) -> bytes:
    if isinstance(source, (str, Path)):
        return Path(source).read_bytes()
    if isinstance(source, bytes):
        return source
    if isinstance(source, bytearray):
        return bytes(source)
    if hasattr(source, "read"):
        can_seek = _safe_seek_zero(source)
        data = source.read()
        if can_seek:
            _safe_seek_zero(source)
        if isinstance(data, str):
            return data.encode("utf-8")
        return bytes(data)
    raise ValueError("Unsupported file source type.")


def _read_csv_headerless_rows(
    buffer: Any,
    is_stream: bool,
    nrows: Optional[int] = None,
) -> pd.DataFrame:
    last_err: Optional[Exception] = None
    for enc in _CSV_FALLBACK_ENCODINGS:
        try:
            rows: List[List[Any]] = []
            if isinstance(buffer, Path):
                with open(buffer, "r", encoding=enc, newline="") as fh:
                    reader = csv.reader(fh, delimiter=",")
                    for row in reader:
                        rows.append([None if cell == "" else cell for cell in row])
                        if nrows is not None and len(rows) >= nrows:
                            break
            elif isinstance(buffer, io.TextIOBase):
                if is_stream:
                    _safe_seek_zero(buffer)
                reader = csv.reader(buffer, delimiter=",")
                for row in reader:
                    rows.append([None if cell == "" else cell for cell in row])
                    if nrows is not None and len(rows) >= nrows:
                        break
            else:
                if is_stream:
                    _safe_seek_zero(buffer)
                wrapper = io.TextIOWrapper(buffer, encoding=enc, newline="")
                try:
                    reader = csv.reader(wrapper, delimiter=",")
                    for row in reader:
                        rows.append([None if cell == "" else cell for cell in row])
                        if nrows is not None and len(rows) >= nrows:
                            break
                finally:
                    try:
                        wrapper.detach()
                    except Exception:
                        pass
            return _rows_to_dataframe(rows, header=None, nrows=nrows)
        except UnicodeDecodeError as exc:
            last_err = exc
    if last_err is not None:
        raise last_err
    return pd.DataFrame()


def _read_csv_with_fallback(
    buffer: Any,
    is_stream: bool,
    header: Optional[int],
    nrows: Optional[int] = None,
) -> pd.DataFrame:
    if header is None:
        return _read_csv_headerless_rows(buffer, is_stream, nrows=nrows)

    last_err: Optional[Exception] = None
    for enc in _CSV_FALLBACK_ENCODINGS:
        try:
            if is_stream:
                _safe_seek_zero(buffer)
            kwargs: dict[str, Any] = {"encoding": enc}
            if header > 0:
                kwargs["skiprows"] = header
                kwargs["header"] = 0
            else:
                kwargs["header"] = 0
            if nrows is not None:
                kwargs["nrows"] = nrows
            return pd.read_csv(buffer, **kwargs)
        except UnicodeDecodeError as exc:
            last_err = exc
    if last_err is not None:
        raise last_err
    if is_stream:
        _safe_seek_zero(buffer)
    kwargs = {"skiprows": header, "header": 0} if header > 0 else {"header": 0}
    if nrows is not None:
        kwargs["nrows"] = nrows
    return pd.read_csv(buffer, **kwargs)


def _decode_rk(rk: int) -> Union[int, float]:
    div100 = bool(rk & 0x01)
    is_int = bool(rk & 0x02)
    if is_int:
        signed_val = struct.unpack("<i", struct.pack("<I", rk & 0xFFFFFFFC))[0] >> 2
        val: Union[int, float] = signed_val
    else:
        val = struct.unpack("<d", b"\x00\x00\x00\x00" + struct.pack("<I", rk & 0xFFFFFFFC))[0]
    if div100:
        val = val / 100.0
    if isinstance(val, float) and val.is_integer() and abs(val) < (1 << 53):
        return int(val)
    return val


def _is_excel_date_format(ifmt: int, fmt_str: str = "") -> bool:
    if ifmt in _BUILTIN_DATE_FORMAT_IDS:
        return True
    if not fmt_str:
        return False
    cleaned = re.sub(r'"[^"]*"', "", fmt_str)
    cleaned = re.sub(r"\[[^\]]*\]", "", cleaned)
    cleaned = re.sub(r"\\.", "", cleaned)
    cleaned_lower = cleaned.lower()
    if any(tok in cleaned_lower for tok in ("yyyy", "yy", "dd", "hh", "mm", "ss", "am/pm", "a/p")):
        if not re.fullmatch(r"[\s0#.,%+Ee/\-()]*", cleaned_lower):
            return True
    return False


def _excel_serial_to_dt(serial: float, datemode: int = 0) -> datetime:
    base = datetime(1904, 1, 1) if datemode == 1 else datetime(1899, 12, 30)
    return base + timedelta(days=serial)


def _decode_biff8_short_string(payload: bytes, offset: int, codepage: str = "cp1252") -> Tuple[str, int]:
    if offset + 2 > len(payload):
        return "", offset
    cch = payload[offset]
    flags = payload[offset + 1]
    pos = offset + 2
    is_unicode = bool(flags & 0x01)
    if is_unicode:
        byte_len = cch * 2
        text = payload[pos:pos + byte_len].decode("utf-16-le", errors="replace")
    else:
        byte_len = cch
        text = payload[pos:pos + byte_len].decode(codepage, errors="replace")
    return text, pos + byte_len


def _decode_biff8_unicode_string(payload: bytes, offset: int, codepage: str = "cp1252") -> Tuple[str, int]:
    if offset + 3 > len(payload):
        return "", offset
    cch = struct.unpack_from("<H", payload, offset)[0]
    flags = payload[offset + 2]
    pos = offset + 3
    is_unicode = bool(flags & 0x01)
    has_phonetic = bool(flags & 0x04)
    has_richtext = bool(flags & 0x08)
    rt_runs = 0
    ext_len = 0
    if has_richtext and pos + 2 <= len(payload):
        rt_runs = struct.unpack_from("<H", payload, pos)[0]
        pos += 2
    if has_phonetic and pos + 4 <= len(payload):
        ext_len = struct.unpack_from("<I", payload, pos)[0]
        pos += 4
    if is_unicode:
        byte_len = cch * 2
        text = payload[pos:pos + byte_len].decode("utf-16-le", errors="replace")
    else:
        byte_len = cch
        text = payload[pos:pos + byte_len].decode(codepage, errors="replace")
    pos += byte_len + (rt_runs * 4) + ext_len
    return text, pos


def _parse_biff8_sst(chunks: List[bytes], codepage: str = "cp1252") -> List[str]:
    if not chunks or len(chunks[0]) < 8:
        return []
    chunk_idx = 0
    pos = 0

    def _read_exact(n: int) -> bytes:
        nonlocal chunk_idx, pos
        out = bytearray()
        while len(out) < n and chunk_idx < len(chunks):
            chunk = chunks[chunk_idx]
            avail = len(chunk) - pos
            if avail <= 0:
                chunk_idx += 1
                pos = 0
                continue
            take = min(n - len(out), avail)
            out.extend(chunk[pos:pos + take])
            pos += take
        return bytes(out)

    def _skip_exact(n: int) -> None:
        nonlocal chunk_idx, pos
        rem = n
        while rem > 0 and chunk_idx < len(chunks):
            avail = len(chunks[chunk_idx]) - pos
            if avail <= 0:
                chunk_idx += 1
                pos = 0
                continue
            take = min(rem, avail)
            pos += take
            rem -= take

    header = _read_exact(8)
    if len(header) < 8:
        return []
    _, cst_unique = struct.unpack("<II", header)
    strings: List[str] = []

    for _ in range(cst_unique):
        hdr = _read_exact(3)
        if len(hdr) < 3:
            break
        cch, flags = struct.unpack("<HB", hdr)
        is_unicode = bool(flags & 0x01)
        has_phonetic = bool(flags & 0x04)
        has_richtext = bool(flags & 0x08)
        rt_runs = struct.unpack("<H", _read_exact(2))[0] if has_richtext else 0
        ext_len = struct.unpack("<I", _read_exact(4))[0] if has_phonetic else 0

        parts: List[str] = []
        chars_left = cch
        while chars_left > 0 and chunk_idx < len(chunks):
            chunk = chunks[chunk_idx]
            if pos >= len(chunk):
                chunk_idx += 1
                pos = 0
                while chunk_idx < len(chunks) and len(chunks[chunk_idx]) == 0:
                    chunk_idx += 1
                if chunk_idx >= len(chunks):
                    break
                cont_flag = chunks[chunk_idx][0]
                pos = 1
                is_unicode = bool(cont_flag & 0x01)
                chunk = chunks[chunk_idx]
            width = 2 if is_unicode else 1
            avail_chars = (len(chunk) - pos) // width
            if avail_chars <= 0:
                chunk_idx += 1
                pos = 0
                while chunk_idx < len(chunks) and len(chunks[chunk_idx]) == 0:
                    chunk_idx += 1
                if chunk_idx >= len(chunks):
                    break
                cont_flag = chunks[chunk_idx][0]
                pos = 1
                is_unicode = bool(cont_flag & 0x01)
                continue
            take_chars = min(chars_left, avail_chars)
            raw_seg = chunk[pos:pos + take_chars * width]
            pos += take_chars * width
            chars_left -= take_chars
            parts.append(raw_seg.decode("utf-16-le" if is_unicode else codepage, errors="replace"))

        if rt_runs:
            _skip_exact(rt_runs * 4)
        if ext_len:
            _skip_exact(ext_len)
        strings.append("".join(parts))

    return strings


def _ole2_extract_workbook_stream(data: bytes) -> bytes:
    if len(data) < 512 or data[:8] != _OLE2_MAGIC:
        raise ValueError("Invalid OLE2 header in .xls file.")

    major_ver = struct.unpack_from("<H", data, 0x1A)[0]
    sec_shift = struct.unpack_from("<H", data, 0x1E)[0]
    mini_shift = struct.unpack_from("<H", data, 0x20)[0]
    if sec_shift not in (9, 12):
        raise ValueError("Unsupported OLE2 sector size in .xls file.")
    sector_size = 1 << sec_shift
    mini_sector_size = 1 << mini_shift if mini_shift <= 12 else 64

    num_fat_sectors = struct.unpack_from("<I", data, 0x2C)[0]
    first_dir_sec = struct.unpack_from("<I", data, 0x30)[0]
    mini_cutoff = struct.unpack_from("<I", data, 0x38)[0]
    first_minifat_sec = struct.unpack_from("<I", data, 0x3C)[0]
    num_minifat_sectors = struct.unpack_from("<I", data, 0x40)[0]
    first_difat_sec = struct.unpack_from("<I", data, 0x44)[0]
    num_difat_sectors = struct.unpack_from("<I", data, 0x48)[0]

    difat: List[int] = list(struct.unpack_from("<109I", data, 0x4C))
    cur_difat = first_difat_sec
    entries_per_sec = sector_size // 4
    seen_difat: Set[int] = set()
    for _ in range(num_difat_sectors):
        if cur_difat >= 0xFFFFFFFA or cur_difat in seen_difat:
            break
        seen_difat.add(cur_difat)
        off = (cur_difat + 1) * sector_size
        if off + sector_size > len(data):
            break
        sec_vals = struct.unpack_from(f"<{entries_per_sec}I", data, off)
        difat.extend(sec_vals[:-1])
        cur_difat = sec_vals[-1]

    fat: List[int] = []
    for fat_sec in difat[:num_fat_sectors]:
        if fat_sec >= 0xFFFFFFFA:
            break
        off = (fat_sec + 1) * sector_size
        if off + sector_size > len(data):
            raise ValueError("Truncated OLE2 FAT sector in .xls file.")
        fat.extend(struct.unpack_from(f"<{entries_per_sec}I", data, off))

    def _read_fat_chain(start_sec: int, max_len: Optional[int] = None) -> bytes:
        out = bytearray()
        sec = start_sec
        visited: Set[int] = set()
        while sec < 0xFFFFFFFA and sec not in visited:
            visited.add(sec)
            off = (sec + 1) * sector_size
            if off + sector_size > len(data):
                break
            out.extend(data[off:off + sector_size])
            if max_len is not None and len(out) >= max_len:
                return bytes(out[:max_len])
            if sec >= len(fat):
                break
            sec = fat[sec]
        return bytes(out[:max_len] if max_len is not None else out)

    dir_bytes = _read_fat_chain(first_dir_sec)
    if len(dir_bytes) < 128:
        raise ValueError("Corrupted OLE2 directory in .xls file.")

    root_entry = None
    target_entry = None
    for off in range(0, len(dir_bytes) - 127, 128):
        entry = dir_bytes[off:off + 128]
        name_len = struct.unpack_from("<H", entry, 0x40)[0]
        obj_type = entry[0x42]
        if obj_type == 0 or name_len < 2 or name_len > 64:
            continue
        name = entry[:name_len - 2].decode("utf-16-le", errors="replace")
        start_sec = struct.unpack_from("<I", entry, 0x74)[0]
        size_low, size_high = struct.unpack_from("<II", entry, 0x78)
        stream_size = size_low if major_ver == 3 else ((size_high << 32) | size_low)
        if obj_type == 5 and root_entry is None:
            root_entry = (name, start_sec, stream_size)
        elif obj_type == 2 and name.lower() in ("workbook", "book"):
            if target_entry is None or name.lower() == "workbook":
                target_entry = (name, start_sec, stream_size)

    if target_entry is None:
        raise ValueError("Could not locate Workbook stream inside .xls OLE2 container.")

    _, wb_start, wb_size = target_entry
    if wb_size == 0:
        raise ValueError("Workbook stream in .xls file is empty.")

    if wb_size < mini_cutoff and root_entry is not None and first_minifat_sec < 0xFFFFFFFA and num_minifat_sectors > 0:
        mini_stream = _read_fat_chain(root_entry[1], root_entry[2])
        minifat_bytes = _read_fat_chain(first_minifat_sec)
        minifat = list(struct.unpack_from(f"<{len(minifat_bytes) // 4}I", minifat_bytes))
        out = bytearray()
        sec = wb_start
        visited: Set[int] = set()
        while sec < 0xFFFFFFFA and sec not in visited and len(out) < wb_size:
            visited.add(sec)
            moff = sec * mini_sector_size
            if moff + mini_sector_size > len(mini_stream):
                break
            out.extend(mini_stream[moff:moff + mini_sector_size])
            if sec >= len(minifat):
                break
            sec = minifat[sec]
        return bytes(out[:wb_size])

    return _read_fat_chain(wb_start, wb_size)


def _dedup_stripped_columns(raw_headers: Any) -> List[str]:
    cols: List[str] = []
    seen_counts: dict[str, int] = {}
    used_cols: set[str] = set()
    for idx, h in enumerate(raw_headers):
        h_str = "" if h is None else str(h).strip()
        base_name = h_str if h_str else f"Unnamed: {idx}"
        if base_name in used_cols:
            suffix = seen_counts.get(base_name, 0) + 1
            candidate = f"{base_name}.{suffix}"
            while candidate in used_cols:
                suffix += 1
                candidate = f"{base_name}.{suffix}"
            seen_counts[base_name] = suffix
            cols.append(candidate)
            used_cols.add(candidate)
        else:
            seen_counts[base_name] = 0
            cols.append(base_name)
            used_cols.add(base_name)
    return cols


def _rows_to_dataframe(
    rows: List[List[Any]],
    header: Optional[int] = 0,
    nrows: Optional[int] = None,
) -> pd.DataFrame:
    if not rows:
        if header is None:
            return pd.DataFrame()
        raise ValueError("The .xls sheet contains no rows.")

    max_cols = max((len(r) for r in rows), default=0)
    if max_cols == 0:
        if header is None:
            return pd.DataFrame()
        raise ValueError("The .xls sheet contains no columns.")

    rect_rows = [list(r) + [None] * (max_cols - len(r)) for r in rows]

    if header is None:
        sliced = rect_rows[:nrows] if nrows is not None else rect_rows
        return pd.DataFrame(sliced)

    if header >= len(rect_rows):
        return pd.DataFrame()

    cols = _dedup_stripped_columns(rect_rows[header])

    data_rows = rect_rows[header + 1:]
    if nrows is not None:
        data_rows = data_rows[:nrows]
    return pd.DataFrame(data_rows, columns=cols)


def _resolve_biff8_codepage(cp_num: int) -> str:
    if cp_num in (1200, 1201):
        return "latin-1"
    if cp_num == 65001:
        return "utf-8"
    if cp_num == 367:
        return "ascii"
    candidate = f"cp{cp_num}"
    try:
        codecs.lookup(candidate)
        return candidate
    except LookupError:
        return "latin-1"


def _read_biff8_xls(
    raw_bytes: bytes,
    sheet_name: Union[str, int] = 0,
    header: Optional[int] = 0,
    nrows: Optional[int] = None,
) -> pd.DataFrame:
    stream = _ole2_extract_workbook_stream(raw_bytes)
    pos = 0
    slen = len(stream)

    def _iter_records(start_pos: int):
        p = start_pos
        while p + 4 <= slen:
            rec_id, rec_len = struct.unpack_from("<HH", stream, p)
            p += 4
            if p + rec_len > slen:
                break
            yield rec_id, stream[p:p + rec_len], p - 4
            p += rec_len

    # Parse workbook globals substream
    codepage = "cp1252"
    datemode = 0
    formats: dict[int, str] = {}
    xf_formats: List[int] = []
    sheets: List[Tuple[int, str]] = []
    sst_chunks: List[bytes] = []
    in_sst = False
    started = False
    bof_depth = 0

    for rec_id, payload, _ in _iter_records(0):
        if not started:
            if rec_id == 0x0809:
                started = True
                bof_depth = 1
            continue
        if rec_id == 0x0809:
            bof_depth += 1
            in_sst = False
            continue
        if rec_id == 0x000A:  # EOF
            in_sst = False
            bof_depth -= 1
            if bof_depth <= 0:
                break
            continue
        if bof_depth > 1:
            continue
        if rec_id == 0x002F:  # FILEPASS
            raise ValueError("Encrypted/password-protected .xls files are not supported.")
        if rec_id != 0x003C:
            in_sst = False
        if rec_id == 0x0042 and len(payload) >= 2:  # CODEPAGE
            cp_num = struct.unpack_from("<H", payload, 0)[0]
            codepage = _resolve_biff8_codepage(cp_num)
        elif rec_id == 0x0022 and len(payload) >= 2:  # DATEMODE
            datemode = struct.unpack_from("<H", payload, 0)[0]
        elif rec_id == 0x041E and len(payload) >= 5:  # FORMAT
            ifmt = struct.unpack_from("<H", payload, 0)[0]
            fmt_str, _ = _decode_biff8_unicode_string(payload, 2, codepage)
            formats[ifmt] = fmt_str
        elif rec_id == 0x00E0 and len(payload) >= 4:  # XF
            ifmt = struct.unpack_from("<H", payload, 2)[0]
            xf_formats.append(ifmt)
        elif rec_id == 0x0085 and len(payload) >= 8:  # BOUNDSHEET
            sheet_pos = struct.unpack_from("<I", payload, 0)[0]
            sname, _ = _decode_biff8_short_string(payload, 6, codepage)
            sheets.append((sheet_pos, sname))
        elif rec_id == 0x00FC:  # SST
            sst_chunks = [payload]
            in_sst = True
        elif rec_id == 0x003C and in_sst:  # CONTINUE for SST
            if payload:
                sst_chunks.append(payload)

    if not sheets:
        raise ValueError("No worksheets found in BIFF8 .xls stream.")

    sst = _parse_biff8_sst(sst_chunks, codepage) if sst_chunks else []

    if isinstance(sheet_name, int):
        if sheet_name < 0 or sheet_name >= len(sheets):
            raise ValueError(f"Worksheet index {sheet_name} out of range for .xls file.")
        target_pos = sheets[sheet_name][0]
    else:
        matched = [s[0] for s in sheets if s[1] == sheet_name]
        if not matched:
            matched = [s[0] for s in sheets if s[1].lower() == str(sheet_name).lower()]
        if not matched:
            raise ValueError(f"Worksheet '{sheet_name}' not found in .xls file.")
        target_pos = matched[0]

    max_row_needed: Optional[int] = None
    if nrows is not None:
        max_row_needed = (nrows - 1) if header is None else (header + nrows)

    cells: dict[Tuple[int, int], Any] = {}
    max_row = -1
    max_col = -1
    ws_started = False
    ws_bof_depth = 0
    pending_formula_cell: Optional[Tuple[int, int]] = None

    def _format_num(num: Union[int, float], xf_idx: int) -> Any:
        if 0 <= xf_idx < len(xf_formats):
            ifmt = xf_formats[xf_idx]
            if _is_excel_date_format(ifmt, formats.get(ifmt, "")):
                try:
                    return _excel_serial_to_dt(float(num), datemode)
                except Exception:
                    return num
        if isinstance(num, float) and num.is_integer() and abs(num) < (1 << 53):
            return int(num)
        return num

    for rec_id, payload, _ in _iter_records(target_pos):
        if not ws_started:
            if rec_id == 0x0809:
                ws_started = True
                ws_bof_depth = 1
            continue
        if rec_id == 0x0809:
            ws_bof_depth += 1
            continue
        if rec_id == 0x000A:  # EOF
            ws_bof_depth -= 1
            if ws_bof_depth <= 0:
                break
            continue
        if ws_bof_depth > 1:
            continue

        if rec_id == 0x0207 and pending_formula_cell is not None:  # STRING for FORMULA
            val_str, _ = _decode_biff8_unicode_string(payload, 0, codepage)
            cells[pending_formula_cell] = val_str
            pending_formula_cell = None
            continue
        if rec_id not in (0x04BC, 0x0221, 0x0236):
            pending_formula_cell = None

        if rec_id == 0x00FD and len(payload) >= 10:  # LABELSST
            r, c, _, isst = struct.unpack_from("<HHHI", payload, 0)
            if max_row_needed is None or r <= max_row_needed:
                cells[(r, c)] = sst[isst] if isst < len(sst) else ""
                max_row = max(max_row, r)
                max_col = max(max_col, c)
        elif rec_id == 0x0204 and len(payload) >= 9:  # LABEL
            r, c, _ = struct.unpack_from("<HHH", payload, 0)
            if max_row_needed is None or r <= max_row_needed:
                val_str, _ = _decode_biff8_unicode_string(payload, 6, codepage)
                cells[(r, c)] = val_str
                max_row = max(max_row, r)
                max_col = max(max_col, c)
        elif rec_id == 0x00D6 and len(payload) >= 9:  # RSTRING
            r, c, _ = struct.unpack_from("<HHH", payload, 0)
            if max_row_needed is None or r <= max_row_needed:
                val_str, _ = _decode_biff8_unicode_string(payload, 6, codepage)
                cells[(r, c)] = val_str
                max_row = max(max_row, r)
                max_col = max(max_col, c)
        elif rec_id == 0x0203 and len(payload) >= 14:  # NUMBER
            r, c, xf, num = struct.unpack_from("<HHHd", payload, 0)
            if max_row_needed is None or r <= max_row_needed:
                cells[(r, c)] = _format_num(num, xf)
                max_row = max(max_row, r)
                max_col = max(max_col, c)
        elif rec_id == 0x027E and len(payload) >= 10:  # RK
            r, c, xf, rk = struct.unpack_from("<HHHI", payload, 0)
            if max_row_needed is None or r <= max_row_needed:
                cells[(r, c)] = _format_num(_decode_rk(rk), xf)
                max_row = max(max_row, r)
                max_col = max(max_col, c)
        elif rec_id == 0x00BD and len(payload) >= 12:  # MULRK
            r, c_first = struct.unpack_from("<HH", payload, 0)
            if max_row_needed is None or r <= max_row_needed:
                n_cols = (len(payload) - 6) // 6
                for i in range(n_cols):
                    xf, rk = struct.unpack_from("<HI", payload, 4 + i * 6)
                    c = c_first + i
                    cells[(r, c)] = _format_num(_decode_rk(rk), xf)
                    max_col = max(max_col, c)
                max_row = max(max_row, r)
        elif rec_id == 0x0205 and len(payload) >= 8:  # BOOLERR
            r, c, _, val_byte, is_err = struct.unpack_from("<HHHBB", payload, 0)
            if max_row_needed is None or r <= max_row_needed:
                cells[(r, c)] = None if is_err else bool(val_byte)
                max_row = max(max_row, r)
                max_col = max(max_col, c)
        elif rec_id == 0x0006 and len(payload) >= 14:  # FORMULA
            r, c, xf = struct.unpack_from("<HHH", payload, 0)
            if max_row_needed is None or r <= max_row_needed:
                res_bytes = payload[6:14]
                if res_bytes[6:8] == b"\xFF\xFF":
                    kind = res_bytes[0]
                    if kind == 0:
                        pending_formula_cell = (r, c)
                    elif kind == 1:
                        cells[(r, c)] = bool(res_bytes[2])
                    elif kind == 3:
                        cells[(r, c)] = ""
                    else:
                        cells[(r, c)] = None
                else:
                    num = struct.unpack_from("<d", res_bytes, 0)[0]
                    cells[(r, c)] = _format_num(num, xf)
                max_row = max(max_row, r)
                max_col = max(max_col, c)

    if max_row < 0 or max_col < 0:
        return _rows_to_dataframe([], header=header, nrows=nrows)

    rows: List[List[Any]] = []
    for r in range(max_row + 1):
        row_vals = [cells.get((r, c), None) for c in range(max_col + 1)]
        rows.append(row_vals)
    return _rows_to_dataframe(rows, header=header, nrows=nrows)


def _xml_local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _xml_attr(elem: ET.Element, target: str) -> Optional[str]:
    for k, v in elem.attrib.items():
        if _xml_local_name(k) == target:
            return v
    return None


def _read_spreadsheetml_xls(
    text: str,
    sheet_name: Union[str, int] = 0,
    header: Optional[int] = 0,
    nrows: Optional[int] = None,
) -> pd.DataFrame:
    root = ET.fromstring(text.lstrip("\ufeff\r\n\t "))
    worksheets = [el for el in root.iter() if _xml_local_name(el.tag) == "Worksheet"]
    if not worksheets:
        raise ValueError("No Worksheet element found in SpreadsheetML .xls file.")

    target_ws = None
    if isinstance(sheet_name, int):
        if 0 <= sheet_name < len(worksheets):
            target_ws = worksheets[sheet_name]
    else:
        for ws in worksheets:
            if (_xml_attr(ws, "Name") or "") == sheet_name:
                target_ws = ws
                break
        if target_ws is None:
            for ws in worksheets:
                if (_xml_attr(ws, "Name") or "").lower() == str(sheet_name).lower():
                    target_ws = ws
                    break
    if target_ws is None:
        raise ValueError(f"Worksheet '{sheet_name}' not found in SpreadsheetML .xls file.")

    table = next((el for el in target_ws if _xml_local_name(el.tag) == "Table"), None)
    if table is None:
        table = next((el for el in target_ws.iter() if _xml_local_name(el.tag) == "Table"), None)
    if table is None:
        return _rows_to_dataframe([], header=header, nrows=nrows)

    rows: List[List[Any]] = []
    current_row_idx = 1
    for child in table:
        if _xml_local_name(child.tag) != "Row":
            continue
        idx_attr = _xml_attr(child, "Index")
        if idx_attr is not None:
            target_row_idx = int(idx_attr)
            while current_row_idx < target_row_idx:
                rows.append([])
                current_row_idx += 1

        row_vals: List[Any] = []
        col_idx = 1
        for cell in child:
            if _xml_local_name(cell.tag) != "Cell":
                continue
            c_idx_attr = _xml_attr(cell, "Index")
            if c_idx_attr is not None:
                target_col = int(c_idx_attr)
                while col_idx < target_col:
                    row_vals.append(None)
                    col_idx += 1

            data_el = next((el for el in cell if _xml_local_name(el.tag) == "Data"), None)
            val: Any = None
            if data_el is not None:
                raw_text = "".join(data_el.itertext())
                dtype = (_xml_attr(data_el, "Type") or "String").lower()
                if dtype == "number":
                    try:
                        fval = float(raw_text)
                        val = int(fval) if fval.is_integer() else fval
                    except ValueError:
                        val = raw_text
                elif dtype == "boolean":
                    val = raw_text.strip() in ("1", "true", "True")
                else:
                    val = raw_text
            row_vals.append(val)
            col_idx += 1

            merge_across = _xml_attr(cell, "MergeAcross")
            if merge_across is not None:
                for _ in range(int(merge_across)):
                    row_vals.append(None)
                    col_idx += 1

        rows.append(row_vals)
        current_row_idx += 1

    return _rows_to_dataframe(rows, header=header, nrows=nrows)


class _HtmlTableParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.table_depth = 0
        self.in_target_table = False
        self.finished_first_table = False
        self.in_cell = False
        self.current_row: List[Any] = []
        self.current_cell_parts: List[str] = []
        self.current_colspan = 1
        self.rows: List[List[Any]] = []

    def _flush_cell(self) -> None:
        if not self.in_cell:
            return
        raw_text = "".join(self.current_cell_parts)
        cleaned = raw_text.replace("\xa0", " ")
        cleaned = re.sub(r"[\r\n\t]+", " ", cleaned).strip()
        self.current_row.append(cleaned if cleaned != "" else None)
        for _ in range(self.current_colspan - 1):
            self.current_row.append(None)
        self.current_cell_parts = []
        self.current_colspan = 1
        self.in_cell = False

    def _flush_row(self) -> None:
        self._flush_cell()
        if self.current_row:
            self.rows.append(self.current_row)
        self.current_row = []

    def handle_starttag(self, tag: str, attrs: List[Tuple[str, Optional[str]]]) -> None:
        t = tag.lower()
        if t == "table":
            if not self.finished_first_table and self.table_depth == 0:
                self.in_target_table = True
            self.table_depth += 1
        elif self.in_target_table and self.table_depth == 1:
            if t == "tr":
                self._flush_row()
            elif t in ("td", "th"):
                self._flush_cell()
                self.in_cell = True
                self.current_cell_parts = []
                self.current_colspan = 1
                for k, v in attrs:
                    if k.lower() == "colspan" and v:
                        try:
                            self.current_colspan = max(1, int(v))
                        except ValueError:
                            self.current_colspan = 1
            elif t == "br" and self.in_cell:
                self.current_cell_parts.append(" ")

    def handle_endtag(self, tag: str) -> None:
        t = tag.lower()
        if t == "table" and self.table_depth > 0:
            if self.table_depth == 1 and self.in_target_table:
                self._flush_row()
            self.table_depth -= 1
            if self.table_depth == 0 and self.in_target_table:
                self.in_target_table = False
                self.finished_first_table = True
        elif self.in_target_table and self.table_depth == 1:
            if t in ("td", "th"):
                self._flush_cell()
            elif t == "tr":
                self._flush_row()

    def handle_data(self, data: str) -> None:
        if self.in_cell:
            self.current_cell_parts.append(data)

    def close(self) -> None:
        super().close()
        if self.in_target_table:
            self._flush_row()
            self.in_target_table = False
            self.finished_first_table = True


def _read_html_xls(
    text: str,
    header: Optional[int] = 0,
    nrows: Optional[int] = None,
) -> pd.DataFrame:
    parser = _HtmlTableParser()
    parser.feed(text)
    parser.close()
    if not parser.rows:
        raise ValueError("No HTML table rows found in .xls file.")
    return _rows_to_dataframe(parser.rows, header=header, nrows=nrows)


def _decode_text_bytes(raw_bytes: bytes) -> str:
    if raw_bytes.startswith((b"\xff\xfe", b"\xfe\xff")):
        return raw_bytes.decode("utf-16")
    if b"\x00" in raw_bytes[:200]:
        even_zeros = sum(1 for b in raw_bytes[:200:2] if b == 0)
        odd_zeros = sum(1 for b in raw_bytes[1:200:2] if b == 0)
        if odd_zeros > len(raw_bytes[:200]) // 4 and even_zeros == 0:
            return raw_bytes.decode("utf-16-le")
        if even_zeros > len(raw_bytes[:200]) // 4 and odd_zeros == 0:
            return raw_bytes.decode("utf-16-be")
    for enc in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
        try:
            return raw_bytes.decode(enc)
        except UnicodeDecodeError:
            continue
    return raw_bytes.decode("latin-1", errors="replace")


def _read_xls_dataframe(
    raw_bytes: bytes,
    sheet_name: Union[str, int] = 0,
    header: Optional[int] = 0,
    nrows: Optional[int] = None,
) -> pd.DataFrame:
    if not raw_bytes:
        raise ValueError("Could not parse .xls file: file is empty.")

    if raw_bytes[:4] == _OOXML_MAGIC:
        kwargs: dict[str, Any] = {"engine": "openpyxl", "sheet_name": sheet_name, "header": header}
        if nrows is not None:
            kwargs["nrows"] = nrows
        return pd.read_excel(io.BytesIO(raw_bytes), **kwargs)

    if raw_bytes[:8] == _OLE2_MAGIC:
        if importlib.util.find_spec("xlrd") is not None:
            kwargs = {"engine": "xlrd", "sheet_name": sheet_name, "header": header}
            if nrows is not None:
                kwargs["nrows"] = nrows
            return pd.read_excel(io.BytesIO(raw_bytes), **kwargs)
        return _read_biff8_xls(raw_bytes, sheet_name=sheet_name, header=header, nrows=nrows)

    text = _decode_text_bytes(raw_bytes)
    head_lower = text.lstrip("\ufeff\r\n\t ").lower()[:4096]

    if ("schemas-microsoft-com:office:spreadsheet" in head_lower) or (
        head_lower.startswith(("<?xml", "<workbook")) and "<worksheet" in head_lower and "<table" in head_lower
    ):
        return _read_spreadsheetml_xls(text, sheet_name=sheet_name, header=header, nrows=nrows)

    if head_lower.startswith(("<!doctype", "<html", "<table", "<head", "<meta", "<style")) or (
        "<table" in head_lower and "<tr" in head_lower
    ):
        return _read_html_xls(text, header=header, nrows=nrows)

    # Guard against arbitrary binary blobs mislabeled as .xls
    sample_ctrl = sum(1 for ch in text[:512] if ord(ch) < 32 and ch not in ("\r", "\n", "\t"))
    if sample_ctrl > 0:
        raise ValueError(
            "Could not parse .xls file: unrecognized or corrupted binary format."
        )

    non_empty_lines = [ln for ln in text.splitlines()[:25] if ln.strip()]
    if not non_empty_lines:
        raise ValueError("Could not parse .xls file: file contains no readable rows.")

    sample = "\n".join(non_empty_lines)
    sep = "\t" if sample.count("\t") >= sample.count(",") and "\t" in sample else None
    if sep is None:
        try:
            dialect = csv.Sniffer().sniff(sample, delimiters=",\t;|")
            sep = dialect.delimiter
        except csv.Error:
            sep = "," if "," in sample else (";" if ";" in sample else None)

    if sep is None:
        raise ValueError(
            "Could not parse .xls file: unrecognized format (expected BIFF8 Excel, SpreadsheetML XML, HTML table, or delimited text)."
        )

    if header is None:
        rows: List[List[Any]] = []
        reader = csv.reader(io.StringIO(text), delimiter=sep)
        for row in reader:
            rows.append([None if cell == "" else cell for cell in row])
            if nrows is not None and len(rows) >= nrows:
                break
        return _rows_to_dataframe(rows, header=None, nrows=nrows)

    csv_kwargs: dict[str, Any] = {"sep": sep}
    if header > 0:
        csv_kwargs["skiprows"] = header
        csv_kwargs["header"] = 0
    else:
        csv_kwargs["header"] = 0
    if nrows is not None:
        csv_kwargs["nrows"] = nrows
    return pd.read_csv(io.StringIO(text), **csv_kwargs)


def detect_header_row(
    file_source: Union[str, bytes, BinaryIO, Path],
    filename: str = "",
    sheet_name: Union[str, int] = 0,
    max_scan_rows: int = 15,
) -> int:
    buffer, is_stream = _get_bytes_or_buffer(file_source)
    name = (filename or (str(file_source) if isinstance(file_source, (str, Path)) else "")).lower()

    try:
        if name.endswith(".csv"):
            df_raw = _read_csv_with_fallback(buffer, is_stream, header=None, nrows=max_scan_rows)
        elif name.endswith(".xls"):
            raw_bytes = _read_raw_bytes(buffer if is_stream else file_source)
            df_raw = _read_xls_dataframe(raw_bytes, sheet_name=sheet_name, header=None, nrows=max_scan_rows)
        else:
            if is_stream:
                _safe_seek_zero(buffer)
            df_raw = pd.read_excel(buffer, header=None, nrows=max_scan_rows, sheet_name=sheet_name)
    except Exception:
        return 1
    finally:
        if is_stream:
            _safe_seek_zero(buffer)

    for row_idx, row in df_raw.iterrows():
        matches = 0
        for val in row.dropna().astype(str):
            v_lower = val.lower().strip()
            for kw in HEADER_KEYWORDS:
                if kw in v_lower:
                    matches += 1
                    break
        if matches >= 2:
            return int(row_idx) + 1

    return 1


def load_dataset(
    source: Union[str, bytes, BinaryIO, Path],
    filename: str = "",
    sheet_name: Union[str, int] = 0,
    header_row: Optional[int] = None,
) -> pd.DataFrame:
    source = _ensure_seekable_source(source)
    name = (filename or (str(source) if isinstance(source, (str, Path)) else "")).lower()
    is_xls = name.endswith(".xls")
    raw_xls_bytes: Optional[bytes] = _read_raw_bytes(source) if is_xls else None

    if header_row is None:
        header_row = detect_header_row(source, filename=filename, sheet_name=sheet_name)

    pandas_header = header_row - 1 if header_row >= 1 else 0
    buffer, is_stream = _get_bytes_or_buffer(source)

    report_date = None
    if pandas_header > 0:
        try:
            if name.endswith(".csv"):
                pre_df = _read_csv_with_fallback(buffer, is_stream, header=None, nrows=pandas_header)
            elif is_xls and raw_xls_bytes is not None:
                pre_df = _read_xls_dataframe(raw_xls_bytes, sheet_name=sheet_name, header=None, nrows=pandas_header)
            else:
                if is_stream:
                    _safe_seek_zero(buffer)
                pre_df = pd.read_excel(buffer, header=None, nrows=pandas_header, sheet_name=sheet_name)
            for _, r in pre_df.iterrows():
                for cell in r.dropna().astype(str):
                    m = re.search(r"(\d{1,2}[-/]\d{1,2}[-/]\d{2,4})", cell)
                    if m:
                        report_date = m.group(1)
                        break
                if report_date:
                    break
        except Exception:
            pass
        finally:
            if is_stream:
                _safe_seek_zero(buffer)

    try:
        if name.endswith(".csv"):
            df = _read_csv_with_fallback(buffer, is_stream, header=pandas_header)
        elif is_xls and raw_xls_bytes is not None:
            df = _read_xls_dataframe(raw_xls_bytes, sheet_name=sheet_name, header=pandas_header)
        else:
            if is_stream:
                _safe_seek_zero(buffer)
            df = pd.read_excel(buffer, header=pandas_header, sheet_name=sheet_name)
    finally:
        if is_stream:
            _safe_seek_zero(buffer)

    df.columns = _dedup_stripped_columns(df.columns)
    df = df.dropna(how="all")
    df.attrs["report_date"] = report_date or datetime.now().strftime("%Y-%m-%d")
    return df
