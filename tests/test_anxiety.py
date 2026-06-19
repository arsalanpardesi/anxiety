"""
Test suite for the anxiety skill.

Stdlib-only (unittest), mirrors the skill's no-dependency promise. Covers the text
extractor, the number/currency parsers, noise stripping, pattern packs, every scanner
check (including the new ones), the diagnostic-only invariant, and CLI behaviour.

Run:
  python3 test_anxiety.py            # or:  python3 -m unittest -v
"""

import base64
import binascii
import json
import os
import subprocess
import sys
import tempfile
import unittest
import zipfile
import zlib
from pathlib import Path

HERE = Path(__file__).resolve().parent
SCRIPTS = HERE.parent / "anxiety" / "scripts"
SCANNER = SCRIPTS / "anxiety_scan.py"
PATTERNS_DIR = HERE.parent / "anxiety" / "patterns"
sys.path.insert(0, str(SCRIPTS))

import anxiety_scan as a            # noqa: E402
import extract_text as et          # noqa: E402
import validate_register as vr     # noqa: E402


# --------------------------------------------------------------------------- helpers

def write(path, text):
    Path(path).write_text(text, encoding="utf-8")
    return str(path)


def make_docx(path, body_xml):
    doc = ('<?xml version="1.0"?><w:document '
           'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
           f'<w:body>{body_xml}</w:body></w:document>')
    with zipfile.ZipFile(path, "w") as z:
        z.writestr("[Content_Types].xml", "<x/>")
        z.writestr("word/document.xml", doc)
    return str(path)


def make_xlsx(path):
    ss = ('<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
          '<si><t>Revenue</t></si></sst>')
    sheet = ('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
             '<sheetData><row><c t="s"><v>0</v></c><c><v>4900000</v></c></row></sheetData></worksheet>')
    with zipfile.ZipFile(path, "w") as z:
        z.writestr("xl/sharedStrings.xml", ss)
        z.writestr("xl/worksheets/sheet1.xml", sheet)
    return str(path)


def make_pptx(path):
    slide = ('<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
             'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
             '<a:t>Slide Title</a:t><a:t>Body content here</a:t></p:sld>')
    with zipfile.ZipFile(path, "w") as z:
        z.writestr("ppt/slides/slide1.xml", slide)
    return str(path)


def make_pdf(path, content_text, compress=False):
    inner = content_text.encode("latin-1")
    content = b"BT /F1 12 Tf 72 720 Td (" + inner + b") Tj ET"
    if compress:
        stream, filt = zlib.compress(content), b"/Filter /FlateDecode "
    else:
        stream, filt = content, b""
    objs = [
        b"<< /Type /Catalog /Pages 2 0 R >>",
        b"<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
        b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Contents 4 0 R >>",
        b"<< /Length %d %b>>\nstream\n%b\nendstream" % (len(stream), filt, stream),
    ]
    out = bytearray(b"%PDF-1.4\n")
    for i, o in enumerate(objs, 1):
        out += b"%d 0 obj\n" % i + o + b"\nendobj\n"
    out += b"trailer\n<< /Size %d /Root 1 0 R >>\n%%%%EOF" % (len(objs) + 1)
    Path(path).write_bytes(out)
    return str(path)


def make_pdf_stream(path, stream, filt_decl):
    """Build a one-page PDF whose content stream bytes and /Filter chain are given verbatim."""
    objs = [
        b"<< /Type /Catalog /Pages 2 0 R >>",
        b"<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
        b"<< /Type /Page /Parent 2 0 R /Contents 4 0 R >>",
        b"<< /Length %d %b>>\nstream\n%b\nendstream" % (len(stream), filt_decl, stream),
    ]
    out = bytearray(b"%PDF-1.4\n")
    for i, o in enumerate(objs, 1):
        out += b"%d 0 obj\n" % i + o + b"\nendobj\n"
    out += b"trailer\n<< /Root 1 0 R >>\n%%%%EOF"
    Path(path).write_bytes(out)
    return str(path)


def pdf_text_obj(text):
    return b"BT /F1 12 Tf 72 720 Td (" + text.encode("latin-1") + b") Tj ET"


def rle_encode(data):
    out = bytearray()
    i = 0
    while i < len(data):
        chunk = data[i:i + 128]
        out.append(len(chunk) - 1)
        out += chunk
        i += 128
    out.append(128)
    return bytes(out)


def lzw_encode(data):
    """PDF LZWDecode-compatible encoder (EarlyChange=1) for round-trip tests."""
    clear, eod = 256, 257
    table = {bytes([i]): i for i in range(256)}
    nextcode = 258
    dsize, dwidth, emitted = 258, 9, 0
    bits = nbits = 0
    out = bytearray()

    def emit(code):
        nonlocal bits, nbits
        bits = (bits << dwidth) | code
        nbits += dwidth
        while nbits >= 8:
            nbits -= 8
            out.append((bits >> nbits) & 0xFF)

    def after(is_data):
        nonlocal dsize, dwidth, emitted
        if is_data:
            emitted += 1
            if emitted >= 2:
                dsize += 1
                if dsize in (511, 1023, 2047) and dwidth < 12:
                    dwidth += 1

    emit(clear)
    after(False)
    w = b""
    for byte in data:
        c = bytes([byte])
        wc = w + c
        if wc in table:
            w = wc
        else:
            emit(table[w]); after(True)
            table[wc] = nextcode; nextcode += 1
            w = c
    if w:
        emit(table[w]); after(True)
    emit(eod)
    if nbits > 0:
        out.append((bits << (8 - nbits)) & 0xFF)
    return bytes(out)


def categories(register):
    return {f["category"] for ph in register["phases"].values() for f in ph["findings"]}


def find(register, category):
    for ph in register["phases"].values():
        for f in ph["findings"]:
            if f["category"] == category:
                return f
    return None


# --------------------------------------------------------------------------- extractor

class TestExtractor(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp()

    def test_plain_text_passthrough(self):
        p = write(os.path.join(self.tmp, "a.md"), "# Title\n\nbody")
        self.assertIn("body", et.load_deliverable_text(p))

    def test_docx_heading_and_table(self):
        body = ('<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Revenue</w:t></w:r></w:p>'
                '<w:p><w:r><w:t>Some prose.</w:t></w:r></w:p>'
                '<w:tbl><w:tr><w:tc><w:p><w:r><w:t>Q1</w:t></w:r></w:p></w:tc>'
                '<w:tc><w:p><w:r><w:t>100</w:t></w:r></w:p></w:tc></w:tr></w:tbl>')
        txt = et.load_deliverable_text(make_docx(os.path.join(self.tmp, "d.docx"), body))
        self.assertIn("# Revenue", txt)
        self.assertIn("| Q1 | 100 |", txt)

    def test_docx_pseudo_heading_bold_and_size(self):
        body = ('<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Executive Summary</w:t></w:r></w:p>'
                '<w:p><w:r><w:rPr><w:sz w:val="36"/></w:rPr><w:t>Market Outlook</w:t></w:r></w:p>'
                '<w:p><w:r><w:t>A long ordinary sentence that should not be treated as a heading.</w:t></w:r></w:p>')
        txt = et.load_deliverable_text(make_docx(os.path.join(self.tmp, "b.docx"), body))
        self.assertIn("### Executive Summary", txt)   # bold -> level 3
        self.assertIn("# Market Outlook", txt)         # 18pt -> level 1
        self.assertNotIn("# A long ordinary", txt)

    def test_xlsx_and_pptx(self):
        self.assertIn("Revenue", et.load_deliverable_text(make_xlsx(os.path.join(self.tmp, "x.xlsx"))))
        self.assertIn("Slide Title", et.load_deliverable_text(make_pptx(os.path.join(self.tmp, "p.pptx"))))

    def test_pdf_raw_and_flate(self):
        raw = make_pdf(os.path.join(self.tmp, "raw.pdf"), "Revenue grew to USD 4.9M.", compress=False)
        flate = make_pdf(os.path.join(self.tmp, "flate.pdf"), "Revenue grew to USD 4.9M.", compress=True)
        self.assertIn("Revenue grew to USD 4.9M", et.load_deliverable_text(raw))
        self.assertIn("Revenue grew to USD 4.9M", et.load_deliverable_text(flate))

    def test_pdf_image_only_rejected(self):
        p = make_pdf(os.path.join(self.tmp, "img.pdf"), "", compress=False)
        # overwrite content with no text-show operators
        Path(p).write_bytes(b"%PDF-1.4\n1 0 obj\n<< /Length 6 >>\nstream\nBT ET\nendstream\nendobj\n%%EOF")
        with self.assertRaises(et.ExtractionError):
            et.load_deliverable_text(p)

    def test_pdf_invalid_rejected(self):
        p = write(os.path.join(self.tmp, "fake.pdf"), "not a pdf at all")
        with self.assertRaises(et.ExtractionError):
            et.load_deliverable_text(p)

    def test_pdf_filter_ascii_hex(self):
        body = pdf_text_obj("Revenue grew to USD 4.9M.")
        p = make_pdf_stream(os.path.join(self.tmp, "hex.pdf"),
                            binascii.hexlify(body) + b">", b"/Filter /ASCIIHexDecode ")
        self.assertIn("Revenue grew to USD 4.9M", et.load_deliverable_text(p))

    def test_pdf_filter_ascii85(self):
        body = pdf_text_obj("Revenue grew to USD 4.9M.")
        p = make_pdf_stream(os.path.join(self.tmp, "a85.pdf"),
                            base64.a85encode(body) + b"~>", b"/Filter /ASCII85Decode ")
        self.assertIn("Revenue grew to USD 4.9M", et.load_deliverable_text(p))

    def test_pdf_filter_run_length(self):
        body = pdf_text_obj("Revenue grew to USD 4.9M.")
        p = make_pdf_stream(os.path.join(self.tmp, "rle.pdf"),
                            rle_encode(body), b"/Filter /RunLengthDecode ")
        self.assertIn("Revenue grew to USD 4.9M", et.load_deliverable_text(p))

    def test_pdf_filter_lzw_small_and_large(self):
        small = pdf_text_obj("Revenue grew to USD 4.9M.")
        p1 = make_pdf_stream(os.path.join(self.tmp, "lzw1.pdf"),
                             lzw_encode(small), b"/Filter /LZWDecode ")
        self.assertIn("Revenue grew to USD 4.9M", et.load_deliverable_text(p1))
        big = pdf_text_obj("Revenue grew across many distinct regions and quarters. " * 60)
        p2 = make_pdf_stream(os.path.join(self.tmp, "lzw2.pdf"),
                             lzw_encode(big), b"/Filter /LZWDecode ")
        self.assertIn("Revenue grew across many distinct regions", et.load_deliverable_text(p2))

    def test_pdf_filter_chain_ascii85_then_flate(self):
        body = pdf_text_obj("Revenue grew to USD 4.9M.")
        stream = base64.a85encode(zlib.compress(body)) + b"~>"
        p = make_pdf_stream(os.path.join(self.tmp, "chain.pdf"),
                            stream, b"/Filter [/ASCII85Decode /FlateDecode] ")
        self.assertIn("Revenue grew to USD 4.9M", et.load_deliverable_text(p))

    def test_pdf_image_filter_skipped(self):
        p = make_pdf_stream(os.path.join(self.tmp, "img.pdf"),
                            b"\xff\xd8\xff\xe0binarydata", b"/Filter /DCTDecode ")
        with self.assertRaises(et.ExtractionError):
            et.load_deliverable_text(p)

    def test_pdf_predictor_stream_skipped(self):
        body = pdf_text_obj("Revenue grew to USD 4.9M.")
        p = make_pdf_stream(os.path.join(self.tmp, "pred.pdf"), zlib.compress(body),
                            b"/Filter /FlateDecode /DecodeParms << /Predictor 12 /Columns 4 >> ")
        with self.assertRaises(et.ExtractionError):
            et.load_deliverable_text(p)

    def test_legacy_format_rejected(self):
        p = write(os.path.join(self.tmp, "old.doc"), "x")
        with self.assertRaises(et.ExtractionError):
            et.load_deliverable_text(p)

    def test_corrupt_office_rejected(self):
        p = write(os.path.join(self.tmp, "fake.docx"), "not a zip")
        with self.assertRaises(et.ExtractionError):
            et.load_deliverable_text(p)


# --------------------------------------------------------------------------- parsers

class TestNumberParsing(unittest.TestCase):
    def test_parse_amount_us_and_eu(self):
        self.assertEqual(a._parse_amount("1,234.56"), 1234.56)
        self.assertEqual(a._parse_amount("1.234,56"), 1234.56)
        self.assertEqual(a._parse_amount("4,9"), 4.9)
        self.assertEqual(a._parse_amount("1,234"), 1234.0)
        self.assertEqual(a._parse_amount("1.234.567"), 1234567.0)
        self.assertIsNone(a._parse_amount("abc"))

    def test_extract_figures_symbols_codes_magnitudes(self):
        def val(s):
            figs = a._extract_labeled_figures(s)
            return figs[0]["value"] if figs else None
        self.assertEqual(val("$4.9M"), 4_900_000)
        self.assertEqual(val("USD 5 billion"), 5_000_000_000)
        self.assertEqual(val("£3.2bn"), 3_200_000_000)
        self.assertEqual(val("€1.234,56"), 1234.56)
        self.assertEqual(val("¥1,200,000"), 1_200_000)
        self.assertEqual(a._extract_labeled_figures("just 42 widgets"), [])  # no currency -> ignored

    def test_currencies_in(self):
        self.assertEqual(a._currencies_in("we made $5M and €3M"), ["EUR", "USD"])
        self.assertEqual(a._currencies_in("only USD 5M and $4M here"), ["USD"])


# --------------------------------------------------------------------------- noise stripping

class TestStripNoise(unittest.TestCase):
    def test_strips_fenced_blockquote_and_example_section(self):
        text = (
            "# Real\n"
            "We conducted interviews with key management.\n\n"
            "```\nfenced: interviews with key management\n```\n\n"
            "> quoted: interviews with key management\n\n"
            "## Example\n"
            "Illustrative: interviews with key management here.\n\n"
            "## Next\n"
            "Back to real content.\n"
        )
        stripped = a._strip_noise(text)
        self.assertEqual(stripped.lower().count("interviews with key management"), 1)
        self.assertIn("Back to real content", stripped)
        self.assertNotIn("fenced:", stripped)
        self.assertNotIn("quoted:", stripped)
        self.assertNotIn("Illustrative:", stripped)

    def test_strips_inline_code_and_quoted_spans(self):
        text = ('The guide warns against writing "we conducted interviews with key management" '
                'or `survey of 50 respondents` without basis.')
        stripped = a._strip_noise(text)
        self.assertNotIn("interviews with key management", stripped)
        self.assertNotIn("survey of 50 respondents", stripped)


# --------------------------------------------------------------------------- pattern packs

class TestPatternPacks(unittest.TestCase):
    def test_en_pack_loads_without_warning(self):
        pat, warn = a.load_patterns("en")
        self.assertIsNone(warn)
        self.assertEqual(pat["language"], "en")
        self.assertIn("baseline_groups", pat)

    def test_missing_language_falls_back(self):
        pat, warn = a.load_patterns("zz")
        self.assertEqual(pat["language"], "en")
        self.assertIsNotNone(warn)

    def test_bad_patterns_file_falls_back(self):
        pat, warn = a.load_patterns("en", "/no/such/pack.json")
        self.assertEqual(pat["language"], "en")
        self.assertIsNotNone(warn)

    def test_shipped_en_json_matches_builtin_keys(self):
        shipped = json.loads((PATTERNS_DIR / "en.json").read_text(encoding="utf-8"))
        for key in a._DEFAULT_PATTERNS:
            self.assertIn(key, shipped, f"en.json missing key: {key}")


# --------------------------------------------------------------------------- scanner checks

class TestScannerChecks(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp()
        self.pat, _ = a.load_patterns("en")

    def reg(self, text, context="", sources=None, advisory=False):
        p = write(os.path.join(self.tmp, "doc.md"), text)
        return a.build_register(p, sources, context, self.pat, advisory_checks=advisory)

    def test_hallucination_critical_when_scope_says_none(self):
        r = self.reg("We conducted interviews with key management about the results.",
                     context="desk research only; no interviews")
        f = find(r, "hallucination")
        self.assertIsNotNone(f)
        self.assertEqual(f["severity"], "critical")

    def test_example_section_not_flagged_as_hallucination(self):
        text = ("# Scope\nDesk research only.\n\n"
                "## Example\nWe conducted interviews with key management (illustrative only).\n")
        r = self.reg(text, context="desk research only")
        self.assertIsNone(find(r, "hallucination"))

    def test_table_body_mismatch(self):
        text = ("Total revenue was $4.9M for the period.\n\n"
                "| Total revenue | $5.4M |\n")
        self.assertIn("table_body_mismatch", categories(self.reg(text)))

    def test_post_cutoff_data(self):
        text = "Results as recorded. In 2025 revenue reached new highs versus 2021."
        r = self.reg(text, context="data through 2022")
        f = find(r, "post_cutoff_data")
        self.assertIsNotNone(f)
        self.assertIn("2025", f["evidence"])

    def test_no_post_cutoff_without_context_year(self):
        text = "In 2025 revenue reached new highs."
        self.assertIsNone(find(self.reg(text), "post_cutoff_data"))

    def test_mixed_currency(self):
        self.assertIn("mixed_currency", categories(self.reg("We earned $5M in the US and €3M in the EU.")))

    def test_mixed_currency_suppressed_with_conversion_note(self):
        text = "We earned $5M in the US and €3M in the EU, converted at the closing exchange rate."
        self.assertNotIn("mixed_currency", categories(self.reg(text)))

    def test_mixed_baseline_same_paragraph(self):
        self.assertIn("mixed_baseline", categories(self.reg("Growth was 12% CAGR though YoY it was 8%.")))

    def test_mixed_baseline_not_flagged_across_paragraphs(self):
        text = "Growth was 12% CAGR over five years.\n\nIn the latest period YoY growth was 8%."
        self.assertNotIn("mixed_baseline", categories(self.reg(text)))

    def test_confidentiality_pii(self):
        text = "Contact admin@example.com. api_key=ABCDEF1234567890. Phone +1 415 555 1234."
        f = find(self.reg(text, advisory=True), "confidentiality_pii")
        self.assertIsNotNone(f)
        self.assertEqual(f["severity"], "high")

    def test_confidentiality_marker_alone_not_flagged(self):
        text = "This document is strictly confidential and summarizes neutral facts."
        self.assertIsNone(find(self.reg(text, advisory=True), "confidentiality_pii"))

    def test_advisory_checks_off_by_default(self):
        text = "Contact admin@example.com about the strictly confidential figures."
        self.assertIsNone(find(self.reg(text), "confidentiality_pii"))
        self.assertIsNotNone(find(self.reg(text, advisory=True), "confidentiality_pii"))

    def test_all_findings_are_not_auto_fixable(self):
        text = ("We conducted interviews with key management.\n\n"
                "Total revenue was $4.9M.\n\n| Total revenue | $5.4M |\n"
                "Contact admin@example.com.")
        r = self.reg(text, context="desk research only")
        all_findings = [f for ph in r["phases"].values() for f in ph["findings"]]
        self.assertTrue(all_findings)
        self.assertTrue(all(f["auto_fixable"] is False for f in all_findings))
        self.assertEqual(r["summary"]["auto_fixable"], 0)

    def test_pdf_end_to_end(self):
        p = make_pdf(os.path.join(self.tmp, "doc.pdf"),
                     "We conducted interviews with key management.", compress=True)
        r = a.build_register(p, None, "desk research only", self.pat)
        self.assertEqual(r["deliverable_format"], "PDF (.pdf)")
        self.assertIsNotNone(find(r, "hallucination"))

    def test_clean_document_scores_high(self):
        r = self.reg("# Intro\n\nThis is a short, clean, neutral statement of fact.\n")
        self.assertGreaterEqual(r["summary"]["overall_confidence"], 90)
        self.assertEqual(r["summary"]["blocking"], 0)
        self.assertIn(r["summary"]["verdict"], ("clean", "review_minor"))

    def test_scoring_blocking_deducts_in_full(self):
        self.assertEqual(a.score_findings([{"severity": "critical"}]), 75)
        self.assertEqual(a.score_findings([{"severity": "high"}, {"severity": "low"}]), 82)

    def test_scoring_advisory_is_capped(self):
        many_medium = [{"severity": "medium"} for _ in range(20)]   # 20*8 = 160 uncapped
        self.assertEqual(a.score_findings(many_medium, a.ADVISORY_CAP_OVERALL),
                         100 - a.ADVISORY_CAP_OVERALL)
        self.assertEqual(a.score_findings(many_medium, a.ADVISORY_CAP_PHASE),
                         100 - a.ADVISORY_CAP_PHASE)

    def test_scoring_ignores_dismissed(self):
        findings = [{"severity": "critical", "status": "dismissed"},
                    {"severity": "low", "status": "open"}]
        self.assertEqual(a.score_findings(findings), 97)
        self.assertEqual(a.verdict_for(findings), "review_minor")

    def test_verdict_bands(self):
        self.assertEqual(a.verdict_for([{"severity": "critical"}]), "not_ready")
        self.assertEqual(a.verdict_for([{"severity": "high"}]), "needs_work")
        self.assertEqual(a.verdict_for([{"severity": "medium"}]), "review_minor")
        self.assertEqual(a.verdict_for([]), "clean")

    def test_summary_has_verdict_and_blocking(self):
        r = self.reg("We conducted interviews with key management.", context="desk research only")
        s = r["summary"]
        self.assertEqual(s["verdict"], "not_ready")
        self.assertIn(s["verdict_label"], a.VERDICT_LABELS.values())
        self.assertGreaterEqual(s["blocking"], 1)


# --------------------------------------------------------------------------- CLI behaviour

class TestCLI(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp()

    def run_cli(self, *args):
        return subprocess.run([sys.executable, str(SCANNER), *args],
                              capture_output=True, text=True)

    def test_banner_and_clean_exit(self):
        p = write(os.path.join(self.tmp, "ok.md"), "# A\n\nNeutral content.\n")
        res = self.run_cli(p, "--out", self.tmp)
        self.assertIn("anxiety", res.stdout.lower())
        self.assertEqual(res.returncode, 0)

    def test_missing_file_exit_code(self):
        res = self.run_cli(os.path.join(self.tmp, "nope.md"))
        self.assertEqual(res.returncode, 2)
        self.assertIn("not found", res.stdout.lower())

    def test_critical_returns_nonzero(self):
        p = write(os.path.join(self.tmp, "c.md"),
                  "Desk research only. We conducted interviews with key management.")
        res = self.run_cli(p, "--context", "desk research only; no interviews", "--out", self.tmp)
        self.assertEqual(res.returncode, 1)

    def test_outputs_written(self):
        p = write(os.path.join(self.tmp, "o.md"), "# A\n\nContent.\n")
        self.run_cli(p, "--out", self.tmp)
        self.assertTrue(os.path.isfile(os.path.join(self.tmp, "o.anxiety.json")))
        self.assertTrue(os.path.isfile(os.path.join(self.tmp, "o.anxiety.md")))


class TestValidator(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp()
        pat, _ = a.load_patterns("en")
        p = write(os.path.join(self.tmp, "d.md"),
                  "We conducted interviews with key management.\n\nContact admin@example.com.")
        self.reg = a.build_register(p, None, "desk research only", pat)

    def test_scanner_register_is_valid(self):
        self.assertEqual(vr.validate_register(self.reg), [])

    def test_detects_bad_severity(self):
        self.reg["phases"]["correctness"]["findings"][0]["severity"] = "extreme"
        errs = vr.validate_register(self.reg)
        self.assertTrue(any("invalid severity" in e for e in errs))

    def test_detects_duplicate_id(self):
        f = self.reg["phases"]["correctness"]["findings"][0]
        dup = dict(f)
        self.reg["phases"]["correctness"]["findings"].append(dup)
        errs = vr.validate_register(self.reg)
        self.assertTrue(any("duplicate finding id" in e for e in errs))

    def test_detects_score_mismatch(self):
        self.reg["summary"]["overall_confidence"] = 999
        errs = vr.validate_register(self.reg)
        self.assertTrue(any("overall_confidence" in e for e in errs))

    def test_detects_missing_key_and_phase_mismatch(self):
        f = self.reg["phases"]["coverage"]["findings"]
        if not f:
            self.reg["phases"]["coverage"]["findings"].append(
                {"id": "X", "phase": "adversarial", "category": "c", "severity": "low",
                 "title": "t", "detail": "d", "evidence": "e", "remediation": "r",
                 "auto_fixable": False})  # missing 'status' + wrong phase
        errs = vr.validate_register(self.reg)
        self.assertTrue(errs)

    def test_cli_round_trip(self):
        path = os.path.join(self.tmp, "r.anxiety.json")
        write(path, json.dumps(self.reg))
        res = subprocess.run([sys.executable, str(SCRIPTS / "validate_register.py"), path],
                             capture_output=True, text=True)
        self.assertEqual(res.returncode, 0)
        self.assertIn("OK", res.stdout)


def make_xlsx_model(path, sheet_cells, name="Model"):
    """sheet_cells: list of (ref, kind, value); kind in {'n','f','e'}."""
    rows = {}
    for ref, kind, value in sheet_cells:
        rnum = int("".join(ch for ch in ref if ch.isdigit()))
        rows.setdefault(rnum, []).append((ref, kind, value))
    body = []
    for rnum in sorted(rows):
        cells = []
        for ref, kind, value in rows[rnum]:
            if kind == "n":
                cells.append(f'<c r="{ref}"><v>{value}</v></c>')
            elif kind == "f":
                cells.append(f'<c r="{ref}"><f>{value}</f><v>0</v></c>')
            elif kind == "e":
                cells.append(f'<c r="{ref}" t="e"><f>1/0</f><v>{value}</v></c>')
        body.append(f'<row r="{rnum}">{"".join(cells)}</row>')
    sheet = ('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
             f'<sheetData>{"".join(body)}</sheetData></worksheet>')
    with zipfile.ZipFile(path, "w") as z:
        z.writestr("xl/worksheets/sheet1.xml", sheet)
    return str(path)


class TestNumericTolerance(unittest.TestCase):
    def test_rounding_not_flagged_as_inconsistent(self):
        # $4.9M vs $4.94M is < 1% apart -> not a conflict.
        self.assertFalse(a._values_conflict({4900000.0, 4940000.0}))

    def test_real_gap_flagged_as_inconsistent(self):
        self.assertTrue(a._values_conflict({4900000.0, 5400000.0}))

    def test_magnitude_restatement_not_conflicting_in_table_body(self):
        # $5M vs $5,000,000 are equal once magnitudes are normalized.
        self.assertFalse(a._table_body_conflict({5000000.0}, {5000000.0}))
        self.assertTrue(a._table_body_conflict({4900000.0}, {5400000.0}))


class TestExcelFormulaChecks(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp()

    def test_read_xlsx_cells(self):
        p = make_xlsx_model(os.path.join(self.tmp, "m.xlsx"),
                            [("A1", "n", 10), ("B1", "f", "A1*2"), ("E1", "e", "#DIV/0!")])
        sheets = et.read_xlsx_cells(p)
        cells = {c["ref"]: c for c in sheets[0]["cells"]}
        self.assertTrue(cells["A1"]["is_number"])
        self.assertTrue(cells["B1"]["has_formula"])
        self.assertEqual(cells["E1"]["error"], "#DIV/0!")

    def test_formula_errors_always_on(self):
        p = make_xlsx_model(os.path.join(self.tmp, "m.xlsx"),
                            [("A1", "n", 10), ("B1", "e", "#REF!")])
        r = a.build_register(p, None, "")  # advisory off
        self.assertIn("formula_errors", categories(r))

    def test_hardcoded_only_with_advisory(self):
        # C1 is a typed number between formula cols B and D.
        cells = [("B1", "f", "A1*2"), ("C1", "n", 99), ("D1", "f", "B1+1")]
        p = make_xlsx_model(os.path.join(self.tmp, "m.xlsx"), cells)
        self.assertNotIn("hardcoded_in_formula_range", categories(a.build_register(p, None, "")))
        self.assertIn("hardcoded_in_formula_range",
                      categories(a.build_register(p, None, "", advisory_checks=True)))


class TestChecklist(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp()
        self.pat, _ = a.load_patterns("en")

    def _write_checklist(self, obj):
        path = os.path.join(self.tmp, "cl.json")
        write(path, json.dumps(obj))
        return path

    def test_checklist_file_loads(self):
        path = self._write_checklist(
            {"name": "Custom", "sections": [
                {"title": "NWC", "severity": "high", "keywords": ["working capital"]}]})
        cl, warn = a.load_checklist(path)
        self.assertIsNone(warn)
        self.assertIn("sections", cl)

    def test_missing_checklist_warns(self):
        cl, warn = a.load_checklist(os.path.join(self.tmp, "nope.json"))
        self.assertIsNone(cl)
        self.assertIsInstance(warn, str)

    def test_malformed_checklist_warns(self):
        path = os.path.join(self.tmp, "bad.json")
        write(path, "{not valid json")
        cl, warn = a.load_checklist(path)
        self.assertIsNone(cl)
        self.assertIsInstance(warn, str)

    def test_missing_section_flagged(self):
        p = write(os.path.join(self.tmp, "q.md"), "# Note\nAdjusted EBITDA after add-backs.\n")
        path = self._write_checklist(
            {"name": "QoE", "sections": [
                {"title": "Net working capital", "severity": "high",
                 "keywords": ["working capital", "nwc"]}]})
        cl, _ = a.load_checklist(path)
        r = a.build_register(p, None, "", self.pat, checklist=cl)
        self.assertIn("missing_expected_section", categories(r))

    def test_present_sections_not_flagged(self):
        # A doc that mentions a section's keyword should not raise it as missing.
        p = write(os.path.join(self.tmp, "q.md"), "Net working capital peg discussion.\n")
        cl = {"name": "t", "sections": [
            {"title": "NWC", "severity": "high", "keywords": ["working capital"]}]}
        r = a.build_register(p, None, "", self.pat, checklist=cl)
        self.assertNotIn("missing_expected_section", categories(r))


class TestCompletenessGate(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp()
        pat, _ = a.load_patterns("en")
        p = write(os.path.join(self.tmp, "d.md"), "# A\n\nContent here.\n")
        self.reg = a.build_register(p, None, "", pat)

    def test_seeds_pending_agent_checks(self):
        self.assertEqual(set(self.reg["agent_checks"]), set(a.AGENT_CHECK_KEYS))
        self.assertTrue(all(v["status"] == "pending"
                            for v in self.reg["agent_checks"].values()))

    def test_require_complete_fails_on_pending(self):
        errs = vr.validate_register(self.reg, require_complete=True)
        self.assertTrue(any("pending" in e for e in errs))

    def test_require_complete_passes_when_done(self):
        for k in self.reg["agent_checks"]:
            self.reg["agent_checks"][k]["status"] = "completed"
        self.assertEqual(vr.validate_register(self.reg, require_complete=True), [])

    def test_not_applicable_counts_as_done(self):
        for k in self.reg["agent_checks"]:
            self.reg["agent_checks"][k]["status"] = "not_applicable"
        self.assertEqual(vr.validate_register(self.reg, require_complete=True), [])

    def test_invalid_status_rejected(self):
        self.reg["agent_checks"]["hallucination"]["status"] = "maybe"
        self.assertTrue(any("invalid or missing status" in e
                            for e in vr.validate_register(self.reg)))


class TestToolBlockAndDiff(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp()
        self.pat, _ = a.load_patterns("en")

    def test_tool_block_present(self):
        p = write(os.path.join(self.tmp, "d.md"), "# A\n\nContent.\n")
        r = a.build_register(p, None, "", self.pat)
        self.assertEqual(r["tool"]["name"], "anxiety")
        self.assertEqual(r["tool"]["version"], a.TOOL_VERSION)
        self.assertFalse(r["tool"]["advisory_checks"])

    def test_diff_reports_resolved_and_new(self):
        base = write(os.path.join(self.tmp, "b.md"),
                     "Revenue was $4.9M.\n\n| Revenue | $5.4M |\n")
        fixed = write(os.path.join(self.tmp, "f.md"),
                      "Revenue was $5.4M.\n\n| Revenue | $5.4M |\n")
        old = a.build_register(base, None, "", self.pat)
        new = a.build_register(fixed, None, "", self.pat)
        d = a.diff_registers(old, new)
        self.assertIn("table_body_mismatch", d["resolved"])
        self.assertEqual(d["new"], [])


class TestMetricKeyingAndGrowthContext(unittest.TestCase):
    """Regression tests for the figure-labeling and extreme-percentage fixes (v0.1.1)."""

    def setUp(self):
        self.tmp = tempfile.mkdtemp()
        self.pat, _ = a.load_patterns("en")

    def reg(self, text, context=""):
        p = write(os.path.join(self.tmp, "doc.md"), text)
        return a.build_register(p, None, context, self.pat)

    def test_prose_vs_prose_revenue_inconsistency_detected(self):
        # Same metric (revenue), worded differently, > 1% apart, both in prose, no table.
        # The old positional 3-word label missed this; canonical metric keys catch it.
        text = ("The target is the highest-growth opportunity in the sector. "
                "Revenue reached $4.9M in 2024.\n\n"
                "## Market Size\nRevenue was $5.4M in 2024 according to our model.\n")
        self.assertIn("inconsistent_figures", categories(self.reg(text)))

    def test_metric_in_following_heading_does_not_leak(self):
        # The 'Market Size' heading after the first figure must not become its label.
        aliases = a._build_metric_aliases(self.pat)
        text = ("Revenue reached $4.9M in 2024.\n\n## Market Size\nRevenue was $5.4M.\n")
        labels = [f["label"] for f in a._extract_labeled_figures(text, aliases)]
        self.assertEqual(labels, ["revenue", "revenue"])

    def test_distinct_metrics_sharing_a_word_not_merged(self):
        # 'revenue growth' and 'revenue' share a word but are different metrics; the longer
        # surface must win so they are not collapsed into a false inconsistency.
        text = ("Revenue was $5.0M for the year.\n\n"
                "Revenue growth was $1.0M in absolute terms.\n")
        self.assertNotIn("inconsistent_figures", categories(self.reg(text)))

    def test_nearest_preceding_metric_wins_in_multi_metric_line(self):
        aliases = a._build_metric_aliases(self.pat)
        labels = [f["label"] for f in
                  a._extract_labeled_figures("Revenue was $5M and EBITDA was $2M.", aliases)]
        self.assertEqual(labels, ["revenue", "ebitda"])

    def test_metric_synonyms_canonicalize(self):
        aliases = a._build_metric_aliases(self.pat)
        figs = a._extract_labeled_figures("Net sales were $5.0M; turnover also hit $5.0M.", aliases)
        self.assertTrue(figs)
        self.assertTrue(all(f["label"] == "revenue" for f in figs))

    def test_growth_percentage_not_flagged_extreme(self):
        self.assertNotIn("extreme_percentages",
                         categories(self.reg("The market grew 250% year over year.")))

    def test_cagr_percentage_not_flagged_extreme(self):
        self.assertNotIn("extreme_percentages",
                         categories(self.reg("Revenue rose at a 180% CAGR over the period.")))

    def test_non_growth_extreme_percentage_still_flagged(self):
        self.assertIn("extreme_percentages",
                      categories(self.reg("The defect ratio hit 480% of the target ceiling.")))


class TestExtractionQuality(unittest.TestCase):
    def test_clean_prose_ok(self):
        ok, reason = et.extraction_quality(
            "This is a perfectly ordinary paragraph of readable English prose. " * 20)
        self.assertTrue(ok)
        self.assertEqual(reason, "")

    def test_short_text_not_judged(self):
        ok, _ = et.extraction_quality("too short to judge")
        self.assertTrue(ok)

    def test_replacement_chars_flagged(self):
        ok, reason = et.extraction_quality(("word " * 60) + ("�" * 80))
        self.assertFalse(ok)
        self.assertIn("replacement", reason)

    def test_lost_spacing_flagged(self):
        ok, reason = et.extraction_quality("Supercalifragilisticexpialidocious" * 30)
        self.assertFalse(ok)

    def test_numeric_spreadsheet_not_flagged(self):
        # A numbers-heavy table is legitimate, not garbled: digits/punctuation must not trip it.
        sheet = "## Sheet: Model\n" + "\n".join(
            "| Revenue | %d | %d | %d |" % (i * 1000, i * 2000, i * 3000) for i in range(40))
        ok, reason = et.extraction_quality(sheet)
        self.assertTrue(ok, reason)

    def test_weird_symbols_flagged(self):
        ok, reason = et.extraction_quality((" word " * 30) + ("\x01\x02\x7fﭖ﷽" * 40))
        self.assertFalse(ok)

    def test_register_carries_extraction_flag(self):
        tmp = tempfile.mkdtemp()
        p = write(os.path.join(tmp, "d.md"),
                  "A clean, readable note with ordinary words and sentences. " * 6)
        r = a.build_register(p, None, "")
        self.assertTrue(r["extraction_ok"])
        self.assertEqual(r["extraction_warning"], "")


if __name__ == "__main__":
    unittest.main(verbosity=2)
