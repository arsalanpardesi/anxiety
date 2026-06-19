"""
Text extraction for the anxiety scan - stdlib only (zipfile + xml.etree + zlib).

Converts the usual knowledge-work file types into normalized markdown-ish text so
the anxiety scanner can interrogate them with one code path:

  .md / .markdown / .txt / .rst   -> read as-is
  .docx                           -> headings (#) + paragraphs + tables (| ... |)
                                     headings come from heading styles, falling back
                                     to bold / large-font cues when no style is set
  .xlsx                           -> one section (##) per sheet, rows as table lines
  .pptx                           -> one section (##) per slide, title as heading
  .pdf                            -> text layer extracted from content streams
                                     (FlateDecode/raw); best-effort, stdlib only

Modern Office formats only (OOXML). Legacy .doc/.xls/.ppt are binary and are not
supported - convert them to the current format first.

PDF extraction is best-effort and depends only on the standard library: it reads the text
layer from content streams, undoing the Flate/LZW/ASCII85/ASCIIHex/RunLength filters (and
chains of them). It cannot read scanned/image-only PDFs (no text layer), encrypted PDFs,
predictor-encoded streams, or custom/CID font encodings - export those to text/markdown
and re-run. When the host harness offers a higher-fidelity reader, prefer it (see SKILL.md).

Usable as a library (`load_deliverable_text(path)`) or a CLI for debugging.

Developed by Arsalan Pardesi. MIT License.
"""

from __future__ import annotations

import re
import sys
import zipfile
import zlib
from pathlib import Path
from typing import Dict, List, Optional, Tuple
from xml.etree import ElementTree as ET

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
S_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

PLAIN_SUFFIXES = {".md", ".markdown", ".txt", ".rst", ".text"}
OFFICE_SUFFIXES = {".docx", ".xlsx", ".pptx"}
PDF_SUFFIXES = {".pdf"}
LEGACY_SUFFIXES = {".doc", ".xls", ".ppt"}

MAX_ROWS_PER_SHEET = 2000


class ExtractionError(Exception):
    pass


def _local(tag: str) -> str:
    return tag.split("}")[-1]


# ---------------------------------------------------------------------------
# DOCX
# ---------------------------------------------------------------------------

def _docx_paragraph_text(p: ET.Element) -> str:
    parts: List[str] = []
    for t in p.iter(f"{{{W_NS}}}t"):
        if t.text:
            parts.append(t.text)
    for br in p.iter(f"{{{W_NS}}}tab"):
        parts.append(" ")
    return "".join(parts).strip()


def _docx_heading_level(p: ET.Element) -> Optional[int]:
    ppr = p.find(f"{{{W_NS}}}pPr")
    if ppr is None:
        return None
    style = ppr.find(f"{{{W_NS}}}pStyle")
    if style is None:
        return None
    val = (style.get(f"{{{W_NS}}}val") or "").lower()
    if "title" in val:
        return 1
    if "heading" in val:
        m = re.search(r"(\d+)", val)
        return min(int(m.group(1)), 6) if m else 2
    return None


def _docx_run_formatting(p: ET.Element) -> tuple:
    """Return (all_runs_bold, max_font_half_points) for a paragraph's text runs."""
    bold_flags: List[bool] = []
    sizes: List[int] = []
    has_text = False
    for r in p.findall(f"{{{W_NS}}}r"):
        if not any(t.text and t.text.strip() for t in r.iter(f"{{{W_NS}}}t")):
            continue
        has_text = True
        rpr = r.find(f"{{{W_NS}}}rPr")
        is_bold = False
        if rpr is not None:
            b = rpr.find(f"{{{W_NS}}}b")
            if b is not None and (b.get(f"{{{W_NS}}}val") or "true").lower() \
                    not in ("0", "false", "off", "none"):
                is_bold = True
            sz = rpr.find(f"{{{W_NS}}}sz")
            if sz is not None:
                try:
                    sizes.append(int(sz.get(f"{{{W_NS}}}val")))
                except (TypeError, ValueError):
                    pass
        bold_flags.append(is_bold)
    all_bold = has_text and all(bold_flags)
    max_sz = max(sizes) if sizes else None
    return all_bold, max_sz


def _docx_pseudo_heading_level(p: ET.Element, text: str) -> Optional[int]:
    """Infer a heading from bold/large-font cues when no heading style is set."""
    words = text.split()
    if not words or len(words) > 14 or len(text) > 120:
        return None
    if text.endswith((".", "!", "?")):
        return None
    all_bold, max_sz = _docx_run_formatting(p)
    if max_sz is not None and max_sz >= 36:   # >= 18pt
        return 1
    if max_sz is not None and max_sz >= 28:   # >= 14pt
        return 2
    if all_bold:
        return 3
    return None


def _docx_table(tbl: ET.Element) -> str:
    rows: List[str] = []
    for tr in tbl.findall(f"{{{W_NS}}}tr"):
        cells = []
        for tc in tr.findall(f"{{{W_NS}}}tc"):
            cell_parts = [_docx_paragraph_text(p) for p in tc.findall(f"{{{W_NS}}}p")]
            cells.append(" ".join(cp for cp in cell_parts if cp).replace("|", "/"))
        if any(c.strip() for c in cells):
            rows.append("| " + " | ".join(cells) + " |")
    return "\n".join(rows)


def _extract_docx(path: str) -> str:
    with zipfile.ZipFile(path) as z:
        xml = z.read("word/document.xml")
    root = ET.fromstring(xml)
    body = root.find(f"{{{W_NS}}}body")
    if body is None:
        return ""
    blocks: List[str] = []
    for el in body:
        tag = _local(el.tag)
        if tag == "p":
            text = _docx_paragraph_text(el)
            if not text:
                continue
            level = _docx_heading_level(el)
            if level is None:
                level = _docx_pseudo_heading_level(el, text)
            blocks.append(("#" * level + " " + text) if level else text)
        elif tag == "tbl":
            tbl = _docx_table(el)
            if tbl:
                blocks.append(tbl)
    return "\n\n".join(blocks)


# ---------------------------------------------------------------------------
# PPTX
# ---------------------------------------------------------------------------

def _slide_sort_key(name: str) -> int:
    m = re.search(r"slide(\d+)\.xml$", name)
    return int(m.group(1)) if m else 0


def _extract_pptx(path: str) -> str:
    blocks: List[str] = []
    with zipfile.ZipFile(path) as z:
        slides = sorted(
            (n for n in z.namelist() if re.match(r"ppt/slides/slide\d+\.xml$", n)),
            key=_slide_sort_key,
        )
        for idx, name in enumerate(slides, 1):
            root = ET.fromstring(z.read(name))
            texts = [t.text.strip() for t in root.iter(f"{{{A_NS}}}t") if t.text and t.text.strip()]
            if not texts:
                continue
            title = texts[0][:120]
            blocks.append(f"## Slide {idx}: {title}")
            body = texts[1:] if len(texts) > 1 else []
            if body:
                blocks.append("\n".join(body))
    return "\n\n".join(blocks)


# ---------------------------------------------------------------------------
# XLSX
# ---------------------------------------------------------------------------

def _xlsx_shared_strings(z: zipfile.ZipFile) -> List[str]:
    if "xl/sharedStrings.xml" not in z.namelist():
        return []
    root = ET.fromstring(z.read("xl/sharedStrings.xml"))
    strings: List[str] = []
    for si in root.findall(f"{{{S_NS}}}si"):
        parts = [t.text for t in si.iter(f"{{{S_NS}}}t") if t.text]
        strings.append("".join(parts))
    return strings


def _xlsx_sheet_map(z: zipfile.ZipFile) -> List[tuple]:
    """Return [(sheet_name, worksheet_path), ...] in workbook order."""
    names = z.namelist()
    if "xl/workbook.xml" not in names:
        sheets = sorted(n for n in names if re.match(r"xl/worksheets/sheet\d+\.xml$", n))
        return [(Path(n).stem, n) for n in sheets]

    wb = ET.fromstring(z.read("xl/workbook.xml"))
    rels_map: Dict[str, str] = {}
    if "xl/_rels/workbook.xml.rels" in names:
        rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        for rel in rels:
            rid = rel.get("Id")
            target = rel.get("Target") or ""
            if rid and target:
                target = target.lstrip("/")
                rels_map[rid] = target if target.startswith("xl/") else f"xl/{target}"

    out: List[tuple] = []
    sheets_el = wb.find(f"{{{S_NS}}}sheets")
    if sheets_el is not None:
        for sheet in sheets_el.findall(f"{{{S_NS}}}sheet"):
            name = sheet.get("name") or "Sheet"
            rid = sheet.get(f"{{{R_NS}}}id")
            path = rels_map.get(rid)
            if path and path in names:
                out.append((name, path))
    if not out:
        sheets = sorted(n for n in names if re.match(r"xl/worksheets/sheet\d+\.xml$", n))
        out = [(Path(n).stem, n) for n in sheets]
    return out


def _xlsx_cell_value(c: ET.Element, shared: List[str]) -> str:
    ctype = c.get("t")
    if ctype == "inlineStr":
        is_el = c.find(f"{{{S_NS}}}is")
        if is_el is not None:
            return "".join(t.text for t in is_el.iter(f"{{{S_NS}}}t") if t.text)
        return ""
    v = c.find(f"{{{S_NS}}}v")
    if v is None or v.text is None:
        return ""
    if ctype == "s":
        try:
            return shared[int(v.text)]
        except (ValueError, IndexError):
            return ""
    return v.text


def _extract_xlsx(path: str) -> str:
    blocks: List[str] = []
    with zipfile.ZipFile(path) as z:
        shared = _xlsx_shared_strings(z)
        for sheet_name, ws_path in _xlsx_sheet_map(z):
            try:
                root = ET.fromstring(z.read(ws_path))
            except KeyError:
                continue
            sheet_data = root.find(f"{{{S_NS}}}sheetData")
            if sheet_data is None:
                continue
            rows: List[str] = []
            truncated = False
            for ri, row in enumerate(sheet_data.findall(f"{{{S_NS}}}row")):
                if ri >= MAX_ROWS_PER_SHEET:
                    truncated = True
                    break
                cells = [
                    _xlsx_cell_value(c, shared).replace("|", "/").replace("\n", " ").strip()
                    for c in row.findall(f"{{{S_NS}}}c")
                ]
                if any(cells):
                    rows.append("| " + " | ".join(cells) + " |")
            if not rows:
                continue
            blocks.append(f"## Sheet: {sheet_name}")
            blocks.append("\n".join(rows))
            if truncated:
                blocks.append(f"_(sheet truncated at {MAX_ROWS_PER_SHEET} rows)_")
    return "\n\n".join(blocks)


def _col_to_index(ref: str) -> int:
    """'B7' -> 2 (1-based column index); 0 if unparseable."""
    m = re.match(r"([A-Za-z]+)", ref or "")
    if not m:
        return 0
    idx = 0
    for ch in m.group(1).upper():
        idx = idx * 26 + (ord(ch) - 64)
    return idx


def read_xlsx_cells(path: str) -> List[Dict]:
    """Return per-sheet cell facts for formula auditing (stdlib only).

    Each sheet: {"name", "cells": [{"ref","col","row","has_formula","error","is_number"}]}.
    Only non-empty cells are returned. Used by the scanner's Excel formula checks.
    """
    sheets: List[Dict] = []
    with zipfile.ZipFile(path) as z:
        for sheet_name, ws_path in _xlsx_sheet_map(z):
            try:
                root = ET.fromstring(z.read(ws_path))
            except KeyError:
                continue
            sheet_data = root.find(f"{{{S_NS}}}sheetData")
            if sheet_data is None:
                continue
            cells: List[Dict] = []
            for ri, row in enumerate(sheet_data.findall(f"{{{S_NS}}}row")):
                if ri >= MAX_ROWS_PER_SHEET:
                    break
                for c in row.findall(f"{{{S_NS}}}c"):
                    ref = c.get("r") or ""
                    ctype = c.get("t")
                    f_el = c.find(f"{{{S_NS}}}f")
                    v_el = c.find(f"{{{S_NS}}}v")
                    has_formula = f_el is not None
                    error = (v_el.text if (ctype == "e" and v_el is not None and v_el.text)
                             else None)
                    is_number = False
                    if not has_formula and ctype in (None, "n") and v_el is not None and v_el.text:
                        try:
                            float(v_el.text)
                            is_number = True
                        except ValueError:
                            is_number = False
                    if not (has_formula or error or is_number):
                        continue
                    rownum = 0
                    mnum = re.search(r"(\d+)$", ref)
                    if mnum:
                        rownum = int(mnum.group(1))
                    cells.append({
                        "ref": ref, "col": _col_to_index(ref), "row": rownum,
                        "has_formula": has_formula, "error": error, "is_number": is_number,
                    })
            if cells:
                sheets.append({"name": sheet_name, "cells": cells})
    return sheets


# ---------------------------------------------------------------------------
# PDF (best-effort, stdlib only)
# ---------------------------------------------------------------------------

_PDF_ESCAPES = {"n": "\n", "r": "\r", "t": "\t", "b": "\b", "f": "\f",
                "(": "(", ")": ")", "\\": "\\"}

# Filter names (and PDF abbreviations) this stdlib reader can undo. Image/binary filters
# (DCTDecode, JPXDecode, CCITTFaxDecode, JBIG2Decode) carry no text layer and are skipped.
_PDF_FILTER_ALIASES = {
    "Fl": "FlateDecode", "AHx": "ASCIIHexDecode", "A85": "ASCII85Decode",
    "LZW": "LZWDecode", "RL": "RunLengthDecode",
}


def _flate_decode(raw: bytes) -> bytes:
    for candidate in (raw, raw.rstrip(b"\r\n"), raw.strip(b"\r\n")):
        for wbits in (zlib.MAX_WBITS, -zlib.MAX_WBITS):
            try:
                return zlib.decompress(candidate, wbits)
            except zlib.error:
                continue
    raise zlib.error("FlateDecode failed")


def _ascii_hex_decode(raw: bytes) -> bytes:
    body = raw.split(b">", 1)[0]
    hx = bytes(c for c in body if not chr(c).isspace())
    if len(hx) % 2:
        hx += b"0"
    return bytes.fromhex(hx.decode("latin-1"))


def _ascii85_decode(raw: bytes) -> bytes:
    body = raw.split(b"~>", 1)[0]
    out = bytearray()
    group: List[int] = []
    for ch in body:
        c = chr(ch)
        if c.isspace():
            continue
        if c == "z" and not group:
            out += b"\x00\x00\x00\x00"
            continue
        if not 33 <= ch <= 117:
            continue
        group.append(ch - 33)
        if len(group) == 5:
            n = 0
            for g in group:
                n = n * 85 + g
            out += n.to_bytes(4, "big")
            group = []
    if group:
        count = len(group)
        group += [84] * (5 - count)
        n = 0
        for g in group:
            n = n * 85 + g
        out += n.to_bytes(4, "big")[: count - 1]
    return bytes(out)


def _lzw_decode(raw: bytes) -> bytes:
    """PDF LZWDecode with the default EarlyChange=1 behaviour (width grows at 511/1023/2047)."""
    clear, eod = 256, 257
    table = [bytes([i]) for i in range(256)] + [b"", b""]
    code_width, prev = 9, None
    out = bytearray()
    bitbuf = bitcnt = i = 0
    n = len(raw)
    while True:
        while bitcnt < code_width and i < n:
            bitbuf = (bitbuf << 8) | raw[i]
            i += 1
            bitcnt += 8
        if bitcnt < code_width:
            break
        bitcnt -= code_width
        code = (bitbuf >> bitcnt) & ((1 << code_width) - 1)
        if code == clear:
            table = [bytes([j]) for j in range(256)] + [b"", b""]
            code_width, prev = 9, None
            continue
        if code == eod:
            break
        if prev is None:
            entry = table[code]
            out += entry
            prev = entry
            continue
        if code < len(table):
            entry = table[code]
        elif code == len(table):
            entry = prev + prev[:1]
        else:
            break
        out += entry
        table.append(prev + entry[:1])
        if len(table) in (511, 1023, 2047) and code_width < 12:
            code_width += 1
        prev = entry
    return bytes(out)


def _run_length_decode(raw: bytes) -> bytes:
    out = bytearray()
    i, n = 0, len(raw)
    while i < n:
        length = raw[i]
        i += 1
        if length == 128:
            break
        if length < 128:
            out += raw[i:i + length + 1]
            i += length + 1
        elif i < n:
            out += bytes([raw[i]]) * (257 - length)
            i += 1
    return bytes(out)


_PDF_FILTERS = {
    "FlateDecode": _flate_decode, "ASCIIHexDecode": _ascii_hex_decode,
    "ASCII85Decode": _ascii85_decode, "LZWDecode": _lzw_decode,
    "RunLengthDecode": _run_length_decode,
}


def _pdf_stream_filters(dict_bytes: bytes) -> List[str]:
    m = re.search(rb"/Filter\s*(\[[^\]]*\]|/[A-Za-z0-9]+)", dict_bytes)
    if not m:
        return []
    names = [n.decode("latin-1") for n in re.findall(rb"/([A-Za-z0-9]+)", m.group(1))]
    return [_PDF_FILTER_ALIASES.get(n, n) for n in names]


def _pdf_decode_streams(data: bytes) -> List[bytes]:
    """Decode every `stream ... endstream` block, honouring its declared /Filter chain."""
    decoded: List[bytes] = []
    for m in re.finditer(rb"stream\r?\n", data):
        start = m.end()
        end = data.find(b"endstream", start)
        if end == -1:
            continue
        raw = data[start:end]
        obj_pos = data.rfind(b"obj", 0, m.start())
        dict_bytes = data[obj_pos:m.start()] if obj_pos != -1 else b""
        # /DecodeParms predictors (used by image/xref streams) are not handled.
        if re.search(rb"/Predictor\s+([2-9]|1[0-9])", dict_bytes):
            continue
        filters = _pdf_stream_filters(dict_bytes)
        if filters:
            if any(f not in _PDF_FILTERS for f in filters):
                continue  # image/binary filter with no text layer
            chunk = raw
            try:
                for f in filters:
                    chunk = _PDF_FILTERS[f](chunk)
            except (zlib.error, ValueError):
                continue
        else:
            try:
                chunk = _flate_decode(raw)
            except zlib.error:
                chunk = raw  # uncompressed content stream
        decoded.append(chunk)
    return decoded


def _pdf_read_literal_string(s: str, j: int) -> Tuple[str, int]:
    """Parse a PDF literal string; s[j-1] == '(' and j points just after it."""
    depth, out, n = 1, [], len(s)
    while j < n and depth > 0:
        c = s[j]
        if c == "\\":
            j += 1
            if j >= n:
                break
            e = s[j]
            if e in _PDF_ESCAPES:
                out.append(_PDF_ESCAPES[e]); j += 1
            elif e == "\n":
                j += 1
            elif e == "\r":
                j += 1
                if j < n and s[j] == "\n":
                    j += 1
            elif e in "01234567":
                digits = e; j += 1; k = 0
                while j < n and k < 2 and s[j] in "01234567":
                    digits += s[j]; j += 1; k += 1
                out.append(chr(int(digits, 8) & 0xFF))
            else:
                out.append(e); j += 1
        elif c == "(":
            depth += 1; out.append(c); j += 1
        elif c == ")":
            depth -= 1
            if depth > 0:
                out.append(c)
            j += 1
        else:
            out.append(c); j += 1
    return "".join(out), j


def _pdf_read_hex_string(s: str, j: int) -> Tuple[str, int]:
    """Parse a PDF hex string; s[j-1] == '<' and j points just after it."""
    n, buf = len(s), []
    while j < n and s[j] != ">":
        if not s[j].isspace():
            buf.append(s[j])
        j += 1
    j += 1
    hx = "".join(buf)
    if len(hx) % 2:
        hx += "0"
    try:
        return bytes.fromhex(hx).decode("latin-1", "replace"), j
    except ValueError:
        return "", j


def _pdf_text_from_content(content: bytes) -> str:
    """Pull the shown text out of one content stream, breaking lines on text-positioning ops."""
    s = content.decode("latin-1", "replace")
    n, i = len(s), 0
    out: List[str] = []
    line: List[str] = []

    def flush() -> None:
        if line:
            out.append("".join(line)); line.clear()

    while i < n:
        c = s[i]
        if c == "(":
            text, i = _pdf_read_literal_string(s, i + 1); line.append(text); continue
        if c == "<" and i + 1 < n and s[i + 1] == "<":
            i += 2; continue
        if c == "<":
            text, i = _pdf_read_hex_string(s, i + 1); line.append(text); continue
        if c == "[":
            i += 1
            while i < n and s[i] != "]":
                if s[i] == "(":
                    text, i = _pdf_read_literal_string(s, i + 1); line.append(text); continue
                if s[i] == "<":
                    text, i = _pdf_read_hex_string(s, i + 1); line.append(text); continue
                i += 1
            i += 1
            continue
        if c.isalpha() or c in "*'\"":
            j = i
            while j < n and (s[j].isalpha() or s[j] in "*'\""):
                j += 1
            op = s[i:j]; i = j
            if op in ("Td", "TD", "T*", "'", '"'):
                flush()
            continue
        i += 1
    flush()
    return "\n".join(seg for seg in out if seg.strip())


def _extract_pdf(path: str) -> str:
    data = Path(path).read_bytes()
    if not data.startswith(b"%PDF"):
        raise ExtractionError(f"'{path}' is not a valid PDF file.")
    chunks: List[str] = []
    for stream in _pdf_decode_streams(data):
        if b"Tj" not in stream and b"TJ" not in stream:
            continue
        text = _pdf_text_from_content(stream)
        if text.strip():
            chunks.append(text)
    result = re.sub(r"\n{3,}", "\n\n", "\n\n".join(chunks)).strip()
    if not result:
        raise ExtractionError(
            "No extractable text layer in PDF (it may be scanned/image-only, encrypted, or use "
            "an unsupported stream filter or font encoding). Export it to text or markdown and re-run."
        )
    return result


# ---------------------------------------------------------------------------
# Dispatch
# ---------------------------------------------------------------------------

_EXTRACTORS = {".docx": _extract_docx, ".xlsx": _extract_xlsx, ".pptx": _extract_pptx}


def load_deliverable_text(path: str) -> str:
    """Return normalized markdown-ish text for any supported deliverable type."""
    p = Path(path)
    suffix = p.suffix.lower()

    if suffix in PLAIN_SUFFIXES or suffix == "":
        return p.read_text(encoding="utf-8", errors="replace")

    if suffix in LEGACY_SUFFIXES:
        raise ExtractionError(
            f"Legacy binary format '{suffix}' is not supported. Save as "
            f"{suffix}x (e.g. .docx/.xlsx/.pptx) and re-run."
        )

    if suffix in PDF_SUFFIXES:
        try:
            return _extract_pdf(path)
        except ExtractionError:
            raise
        except Exception as exc:  # noqa: BLE001 - surface a clean message
            raise ExtractionError(f"Failed to extract PDF '{path}': {exc}") from exc

    if suffix in _EXTRACTORS:
        if not zipfile.is_zipfile(path):
            raise ExtractionError(f"'{path}' is not a valid {suffix} (OOXML) file.")
        try:
            return _EXTRACTORS[suffix](path)
        except ExtractionError:
            raise
        except Exception as exc:  # noqa: BLE001 - surface a clean message
            raise ExtractionError(f"Failed to extract {suffix} '{path}': {exc}") from exc

    # Unknown extension: best-effort plain-text read.
    return p.read_text(encoding="utf-8", errors="replace")


def main() -> int:
    if len(sys.argv) < 2:
        print("Usage: python extract_text.py <file> [out.md]")
        return 2
    try:
        text = load_deliverable_text(sys.argv[1])
    except ExtractionError as exc:
        print(f"Error: {exc}")
        return 1
    if len(sys.argv) >= 3:
        Path(sys.argv[2]).write_text(text, encoding="utf-8")
        print(f"Wrote {len(text)} chars to {sys.argv[2]}")
    else:
        print(text)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
