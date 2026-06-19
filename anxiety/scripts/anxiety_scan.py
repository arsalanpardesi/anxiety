"""
Anxiety scan - re-interrogate any completed knowledge-work deliverable.

Three phases, domain-agnostic, stdlib-only (no third-party deps, no DB, no network):
  Phase 1 (Coverage)    - what was missed?
  Phase 2 (Correctness) - is what was said actually right?
  Phase 3 (Adversarial) - what would a hostile reader attack?

Outputs a JSON register and a markdown summary. Diagnostic only: it never edits
the deliverable.

The language-specific cues (hedges, deferred-analysis phrases, hallucination
patterns, direction words, etc.) live in pattern packs under ../patterns. English
ships built-in and as patterns/en.json; add a sibling JSON for another language and
select it with --lang. The scanner always falls back to the built-in English pack if
a pack is missing or unreadable, so it never hard-fails.

Usage:
  python anxiety_scan.py <deliverable.md> [--sources <dir>] [--context "<scope note>"]
                         [--lang <code>] [--patterns <file.json>] [--out <dir>]

Developed by Arsalan Pardesi. MIT License.
"""

from __future__ import annotations

import argparse
import json
import os
import re
import sys
import uuid
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from extract_text import (ExtractionError, extraction_quality,  # noqa: E402
                          load_deliverable_text, read_xlsx_cells)

CRITICAL, HIGH, MEDIUM, LOW = "critical", "high", "medium", "low"
_SEVERITY_ORDER = {CRITICAL: 0, HIGH: 1, MEDIUM: 2, LOW: 3}
_PENALTY = {CRITICAL: 25, HIGH: 15, MEDIUM: 8, LOW: 3}


def _finding(fid, phase, category, severity, title, detail,
             evidence="", remediation="", auto_fixable=False) -> Dict[str, Any]:
    return {
        "id": fid, "phase": phase, "category": category, "severity": severity,
        "title": title, "detail": detail, "evidence": evidence,
        "remediation": remediation, "auto_fixable": auto_fixable, "status": "open",
    }


# ---------------------------------------------------------------------------
# Pattern packs (language-specific cues, externalized so other languages can be
# added without touching code). The built-in English pack below is the fallback.
# ---------------------------------------------------------------------------

_PATTERNS_DIR = Path(__file__).resolve().parent.parent / "patterns"

_DEFAULT_PATTERNS: Dict[str, Any] = {
    "language": "en",
    "stopwords": [
        "the", "and", "for", "with", "from", "this", "that", "into", "your", "their",
        "report", "memo", "draft", "final", "notes", "copy", "version", "document",
        "analysis", "summary", "overview", "data", "file", "sheet", "page", "appendix",
    ],
    "direction_positive": [
        r"\bincreas(?:ed|ing)\b", r"\bgr[eo]w(?:th|n|ing)?\b", r"\bimproved?\b",
        r"\brose\b", r"\brising\b", r"\bhigher\b", r"\bup\b",
    ],
    "direction_negative": [
        r"\bdeclin(?:ed|ing|e)\b", r"\bfell\b", r"\breduced?\b", r"\bdropped?\b",
        r"\bdecreas(?:ed|ing)\b", r"\bdeteriorat(?:ed|ing)\b", r"\blower(?:ed)?\b",
    ],
    "hallucination_patterns": [
        [r"(?:\d+|one|two|three|four|five|several|multiple)\s+interviews?\s+(?:were|was|conducted)", "claims interviews were conducted"],
        [r"(?:we|our team|the team)\s+(?:conducted|performed|held)\s+(?:a\s+)?(?:site\s+visit|interview|workshop|survey)", "claims a visit/interview/workshop/survey was conducted"],
        [r"management\s+(?:confirmed|stated|indicated)\s+(?:during|in)\s+(?:our|the)\s+(?:session|call|meeting|interview)", "claims management confirmed during a session"],
        [r"(?:expert|management|stakeholder)\s+(?:call|session|interview)\s+(?:on|dated|of)\s+\d", "references a specific dated call/session"],
        [r"(?:interviews?|discussions?|conversations?)\s+with\s+(?:key\s+)?(?:management|personnel|staff|stakeholders|participants)", "claims interviews/discussions with people"],
        [r"(?:physical|on-?site)\s+(?:inspection|assessment|review|observation)\s+(?:revealed|confirmed|showed|found)", "claims on-site inspection findings"],
        [r"(?:we\s+)?(?:observed|inspected|toured|visited)\s+(?:the\s+)?(?:plant|facility|site|office|premises)", "claims direct observation of a facility"],
        [r"(?:our\s+)?testing\s+(?:confirmed|revealed|showed|demonstrated)", "claims testing was performed"],
        [r"(?:a\s+)?survey\s+of\s+\d+\s+(?:respondents?|participants?|customers?|users?)", "claims a quantified survey"],
    ],
    "scope_none_pattern": (
        r"(?:desk\s+research\s+only|no\s+(?:interviews?|management\s+access|site\s+visits?|"
        r"primary\s+research)|access\s*[:\-]\s*none|secondary\s+sources\s+only)"
    ),
    "hedges": [
        r"\bbroadly\b", r"\bapproximately\b", r"\blargely\b", r"\bessentially\b",
        r"\broughly\b", r"\bgenerally\b", r"\btypically\b", r"\bsubstantially\b",
        r"\bmostly\b", r"\brelatively\b", r"\bsomewhat\b", r"\bfairly\b",
    ],
    "deferred": [
        [r"should\s+be\s+analy[sz]ed", "should be analyzed"],
        [r"should\s+be\s+reviewed", "should be reviewed"],
        [r"requires?\s+further\s+(?:analysis|investigation|review|work)", "requires further analysis"],
        [r"could\s+not\s+(?:be\s+)?(?:confirmed|verified|validated)", "could not be confirmed"],
        [r"was\s+not\s+(?:available|provided|disclosed)", "was not available"],
        [r"(?:insufficient|limited)\s+(?:data|evidence|information)", "insufficient data"],
        [r"to\s+be\s+(?:determined|confirmed)|\bTBD\b|\bTBC\b", "to be determined"],
    ],
    "citation_hint": (
        r"(?:source\s*[:\-]|ref\s*[:\-]|cite|citation|per\s+the|according\s+to|see\s+(?:table|figure|"
        r"appendix|section|exhibit)|\[\d+\]|\(\d{4}\)|et\s+al\.|https?://|footnote|endnote)"
    ),
    "superlative_pattern": (
        r"\b(?:the\s+)?(?:most|least|best|worst|highest|lowest|largest|"
        r"smallest|strongest|weakest|fastest|slowest|leading)\b"
    ),
    "superlative_qualifier_pattern": (
        r"benchmark|peer|comparator|sector|industry|quartile|median|percentile|"
        r"compared\s+to|relative\s+to|versus|vs\.?"
    ),
    "baseline_groups": {
        "growth basis": [r"\bCAGR\b", r"\byear[- ]over[- ]year\b|\bYoY\b"],
        "value basis": [r"\bnominal\b", r"\breal\s+terms?\b|\bin\s+real\b"],
        "period basis": [r"\bper\s+annum\b|\bannual(?:ly|i[sz]ed)?\b", r"\bper\s+month\b|\bmonthly\b"],
    },
    # A percentage > 100% is normal inside a growth/CAGR context, so the extreme-percentage
    # check is suppressed when any of these cues sit near the value.
    "growth_context": (
        r"\bgr(?:ew|ow(?:s|n|ing|th)?)\b|\bincreas(?:e|ed|ing)\b|\brose\b|\brising\b|"
        r"\bCAGR\b|\bcompounded?\b|\byear[- ]over[- ]year\b|\bYoY\b|\bY/Y\b|\bmultiple\b"
    ),
    # Canonical metric names -> the surface phrases that denote them. Used to key figures by
    # the metric they describe (order-independent), so the same metric stated two different
    # ways is recognized as one quantity. Longest matching surface in a figure's sentence wins,
    # so "revenue growth" / "gross margin" don't collapse into "revenue" / "margin".
    "metric_aliases": {
        "revenue": ["revenue", "revenues", "net sales", "sales", "turnover", "top line"],
        "revenue growth": ["revenue growth", "sales growth"],
        "gross profit": ["gross profit"],
        "gross margin": ["gross margin"],
        "operating margin": ["operating margin"],
        "ebitda margin": ["ebitda margin"],
        "net margin": ["net margin"],
        "ebitda": ["adjusted ebitda", "ebitda"],
        "ebit": ["operating income", "operating profit", "ebit"],
        "net income": ["net income", "net profit", "net earnings", "bottom line"],
        "net debt": ["net debt"],
        "gross debt": ["gross debt", "total debt"],
        "enterprise value": ["enterprise value", "ev"],
        "equity value": ["equity value", "market capitalization", "market cap"],
        "capex": ["capital expenditures", "capital expenditure", "capex"],
        "free cash flow": ["free cash flow", "fcf"],
        "working capital": ["net working capital", "working capital", "nwc"],
        "headcount": ["headcount", "employees", "fte"],
        "arr": ["annual recurring revenue", "arr"],
        "churn": ["churn rate", "churn"],
        "market size": ["addressable market", "market size", "tam"],
        "cagr": ["cagr"],
    },
}

# Updated by load_patterns(); _significant_words() reads it.
STOPWORDS = set(_DEFAULT_PATTERNS["stopwords"])


def load_patterns(lang: str = "en",
                  patterns_file: Optional[str] = None) -> Tuple[Dict[str, Any], Optional[str]]:
    """Return (patterns, warning). Falls back to built-in English if a pack is missing."""
    global STOPWORDS
    pat: Dict[str, Any] = json.loads(json.dumps(_DEFAULT_PATTERNS))  # deep copy

    src: Optional[Path] = None
    if patterns_file:
        src = Path(patterns_file)
    elif lang:
        candidate = _PATTERNS_DIR / f"{lang}.json"
        # Only bother loading a file when it differs from the built-in default.
        if candidate.is_file():
            src = candidate

    warning: Optional[str] = None
    loaded = False
    if src is not None:
        if src.is_file():
            try:
                data = json.loads(src.read_text(encoding="utf-8"))
                if not isinstance(data, dict):
                    raise ValueError("pattern pack must be a JSON object")
                pat.update(data)
                loaded = True
            except Exception as exc:  # noqa: BLE001 - surface a clean warning
                warning = (f"Could not read pattern pack '{src}': {exc}. "
                           "Using built-in English patterns.")
        elif patterns_file:
            warning = (f"Pattern pack not found: '{src}'. Using built-in English patterns.")

    if patterns_file and not loaded and warning is None:
        warning = f"Pattern pack not found: '{patterns_file}'. Using built-in English patterns."
    if lang and lang != "en" and not loaded and warning is None:
        warning = (f"No pattern pack for language '{lang}' at {_PATTERNS_DIR / (lang + '.json')}. "
                   "Using built-in English patterns.")

    pat["language"] = pat.get("language", lang) if loaded else "en"
    STOPWORDS = set(pat.get("stopwords") or [])
    return pat, warning


def _significant_words(text: str, min_len: int = 4) -> List[str]:
    return [w for w in re.split(r"\W+", text.lower())
            if len(w) >= min_len and w not in STOPWORDS]


# ---------------------------------------------------------------------------
# Section parsing (generic markdown headings)
# ---------------------------------------------------------------------------

def _parse_sections(text: str) -> List[Dict[str, Any]]:
    lines = text.split("\n")
    sections: List[Dict[str, Any]] = []
    heading_re = re.compile(r"^(#{1,6})\s+(.+?)\s*#*$")
    open_sec: Optional[Dict[str, Any]] = None
    for i, line in enumerate(lines):
        m = heading_re.match(line.strip())
        if m:
            if open_sec is not None:
                open_sec["body"] = "\n".join(lines[open_sec["start"] + 1:i])
                sections.append(open_sec)
            open_sec = {"level": len(m.group(1)), "title": m.group(2).strip(),
                        "start": i, "body": ""}
    if open_sec is not None:
        open_sec["body"] = "\n".join(lines[open_sec["start"] + 1:])
        sections.append(open_sec)
    return sections


def _substantive_paragraphs(body: str) -> List[str]:
    paras = re.split(r"\n\s*\n", body)
    out = []
    for p in paras:
        s = p.strip()
        if len(s) < 60:
            continue
        if s.startswith("|") or s.startswith("#") or s.startswith("```"):
            continue
        out.append(s)
    return out


_EXAMPLE_HEADING_RE = re.compile(r"\bexamples?\b|\bsample\b|\billustrative\b|\btemplate\b|\bspecimen\b",
                                 re.IGNORECASE)


def _strip_noise(text: str) -> str:
    """Remove material that is quoted or illustrative rather than the author's own claim, so
    the pattern-based checks do not mistake specimens for real assertions. Removes:
      - fenced code blocks (``` ... ``` / ~~~ ... ~~~)
      - inline code spans (`...`)
      - inline quoted spans ("..." and curly “...”) kept on a single line
      - block-quote lines (> ...)
      - whole sections whose heading matches example/sample/illustrative/template/specimen
    """
    text = re.sub(r"```.*?```", " ", text, flags=re.DOTALL)
    text = re.sub(r"~~~.*?~~~", " ", text, flags=re.DOTALL)
    text = re.sub(r"`[^`\n]{1,200}`", " ", text)
    text = re.sub(r"\"[^\"\n]{1,200}\"", " ", text)
    text = re.sub(r"\u201c[^\u201d\n]{1,200}\u201d", " ", text)
    heading_re = re.compile(r"^(#{1,6})\s+(.+?)\s*#*$")
    out: List[str] = []
    skip_level: Optional[int] = None
    for line in text.split("\n"):
        st = line.strip()
        m = heading_re.match(st)
        if m:
            level = len(m.group(1))
            if skip_level is not None and level <= skip_level:
                skip_level = None
            if skip_level is None and _EXAMPLE_HEADING_RE.search(m.group(2)):
                skip_level = level
                continue
        if skip_level is not None:
            continue
        if st.startswith(">"):
            continue
        out.append(line)
    return "\n".join(out)


# ---------------------------------------------------------------------------
# Phase 1 - Coverage
# ---------------------------------------------------------------------------

def _phase1_coverage(text: str, sources_dir: Optional[str]) -> List[Dict]:
    findings: List[Dict] = []
    text_lower = text.lower()
    n = 0

    if sources_dir and os.path.isdir(sources_dir):
        unreferenced = []
        for root, _dirs, files in os.walk(sources_dir):
            for fname in files:
                if fname.startswith("."):
                    continue
                stem = Path(fname).stem
                words = _significant_words(stem)
                if not words:
                    continue
                threshold = min(2, len(words))
                matches = sum(1 for w in words if w in text_lower)
                if matches < threshold:
                    rel = os.path.relpath(os.path.join(root, fname), sources_dir)
                    unreferenced.append(rel)
        if unreferenced:
            sev = CRITICAL if len(unreferenced) > 10 else HIGH if len(unreferenced) > 5 else MEDIUM
            n += 1
            findings.append(_finding(
                f"COV-{n:03d}", "coverage", "unreferenced_sources", sev,
                f"{len(unreferenced)} source file(s) not referenced in the deliverable",
                "These files in the sources directory have a subject that never appears in "
                "the deliverable. Each may hold evidence that was overlooked.",
                evidence="\n".join(f"  - {u}" for u in unreferenced[:20])
                + (f"\n  ... and {len(unreferenced) - 20} more" if len(unreferenced) > 20 else ""),
                remediation="Open each unreferenced source. Where relevant, fold its findings "
                "into the deliverable; otherwise note explicitly why it is out of scope.",
            ))

    # Thin sections
    thin = []
    for sec in _parse_sections(text):
        if sec["level"] > 3:
            continue
        paras = _substantive_paragraphs(sec["body"])
        if len(paras) < 2 and not re.search(r"^\s*\|", sec["body"], re.MULTILINE):
            thin.append(f"{sec['title'][:70]} ({len(paras)} substantive paragraph(s))")
    if thin:
        sev = HIGH if len(thin) > 3 else MEDIUM
        n += 1
        findings.append(_finding(
            f"COV-{n:03d}", "coverage", "thin_sections", sev,
            f"{len(thin)} section(s) have insufficient depth",
            "These sections contain very little substantive content. A reviewer would question "
            "whether the topic received adequate attention.",
            evidence="\n".join(f"  - {t}" for t in thin[:15]),
            remediation="Expand each thin section with real analysis, or merge/remove it if the "
            "topic does not warrant a standalone section.",
        ))

    return findings


# ---------------------------------------------------------------------------
# Phase 2 - Correctness
# ---------------------------------------------------------------------------

# Monetary figure parsing. Recognizes a leading or trailing currency marker
# (symbol or ISO code), magnitude words (k/m/bn/trillion...), and both US
# (1,234.56) and European (1.234,56 / 1 234,56) number formats.
_CUR_SYMBOLS = "£$€¥₹₩₽₺₪฿"
_ISO_CODES = [
    "USD", "EUR", "GBP", "JPY", "CNY", "RMB", "INR", "AUD", "CAD", "CHF", "HKD",
    "SGD", "KRW", "BRL", "RUB", "ZAR", "MXN", "SEK", "NOK", "DKK", "PLN", "AED",
    "SAR", "TRY", "THB", "IDR", "MYR", "PHP", "NZD", "ILS",
]
_MAGNITUDES = {
    "k": 1e3, "thousand": 1e3,
    "m": 1e6, "mm": 1e6, "mn": 1e6, "million": 1e6,
    "b": 1e9, "bn": 1e9, "billion": 1e9,
    "t": 1e12, "tn": 1e12, "trillion": 1e12,
}

_SYM_RE = "[" + re.escape(_CUR_SYMBOLS) + "]"
_ISO_RE = "(?:" + "|".join(_ISO_CODES) + ")"
_NUM_RE = r"\d{1,3}(?:[.,\s]\d{3})+(?:[.,]\d+)?|\d+(?:[.,]\d+)?|\d+"
_MAG_RE = r"(?:thousand|million|billion|trillion|mm|mn|bn|tn|k|m|b|t)(?![A-Za-z])"

_FIGURE_RE = re.compile(
    rf"(?:(?P<cur1>{_SYM_RE}|\b{_ISO_RE}\b)\s*\(?\s*(?P<n1>{_NUM_RE})\s*(?P<m1>{_MAG_RE})?)"
    rf"|(?:(?P<n2>{_NUM_RE})\s*(?P<m2>{_MAG_RE})?\s*(?P<cur2>{_SYM_RE}|\b{_ISO_RE}\b))",
    re.IGNORECASE,
)

# Normalize symbols to ISO codes so "$" and "USD" are not counted as two currencies.
_SYMBOL_TO_CODE = {
    "$": "USD", "£": "GBP", "€": "EUR", "¥": "JPY", "₹": "INR",
    "₩": "KRW", "₽": "RUB", "₺": "TRY", "₪": "ILS", "฿": "THB",
}

# Confidentiality / PII signals (language-independent).
_PII_PATTERNS = [
    (r"\b[\w.+-]+@[\w-]+\.[\w.-]{2,}\b", "email address", HIGH),
    (r"(?:\+\d{1,3}[\s.\-]?\d{2,4}[\s.\-]?\d{3,4}[\s.\-]?\d{2,4})|(?:\(\d{3}\)\s?\d{3}[\s.\-]\d{4})",
     "phone-like number", MEDIUM),
    (r"\b(?:AKIA|ASIA)[0-9A-Z]{16}\b", "AWS access key", HIGH),
    (r"\bsk-[A-Za-z0-9]{20,}\b", "API token", HIGH),
    (r"-----BEGIN (?:RSA |EC |OPENSSH )?PRIVATE KEY-----", "private key block", HIGH),
    (r"(?i)\b(?:api[_-]?key|secret|password|passwd|client[_-]?secret|access[_-]?token)\b\s*[:=]\s*\S+",
     "credential assignment", HIGH),
    (r"\b\d{3}-\d{2}-\d{4}\b", "SSN-like number", HIGH),
    (r"(?i)\b(?:strictly\s+)?confidential\b|\bdo\s+not\s+distribute\b|\bnot\s+for\s+distribution\b|"
     r"\bprivileged\s+(?:and|&)\s+confidential\b", "confidentiality marker", LOW),
]


def _currencies_in(text: str) -> List[str]:
    """Distinct ISO currency codes attached to monetary figures in the text."""
    codes = set()
    for m in _FIGURE_RE.finditer(text):
        cur = (m.group("cur1") or m.group("cur2") or "").strip().upper()
        if not cur:
            continue
        codes.add(_SYMBOL_TO_CODE.get(cur, cur) if len(cur) == 1 else cur)
    return sorted(codes)


def _parse_amount(s: str) -> Optional[float]:
    """Parse a numeric string in US or European convention into a float."""
    s = s.strip().replace(" ", "")
    if not s:
        return None
    has_comma, has_dot = "," in s, "." in s
    try:
        if has_comma and has_dot:
            if s.rfind(",") > s.rfind("."):       # 1.234,56 -> comma is decimal
                s = s.replace(".", "").replace(",", ".")
            else:                                  # 1,234.56 -> comma is thousands
                s = s.replace(",", "")
        elif has_comma:
            parts = s.split(",")
            if len(parts) == 2 and len(parts[1]) != 3:   # 4,9 -> decimal
                s = s.replace(",", ".")
            else:                                          # 1,234 / 1,234,567 -> thousands
                s = s.replace(",", "")
        elif has_dot:
            if s.count(".") > 1:                   # 1.234.567 -> thousands
                s = s.replace(".", "")
            # single dot: keep as decimal (US convention)
        return float(s)
    except ValueError:
        return None


def _build_metric_aliases(pat: Dict[str, Any]) -> List[Tuple[str, Any]]:
    """Compile (canonical_name, surface_regex) pairs from the pattern pack's metric_aliases."""
    out: List[Tuple[str, Any]] = []
    for canon, surfaces in (pat.get("metric_aliases") or {}).items():
        for s in surfaces:
            if s:
                out.append((canon, re.compile(r"\b" + re.escape(str(s)) + r"\b", re.IGNORECASE)))
    return out


def _metric_key(text: str, fig_start: int, fig_end: int,
                aliases: List[Tuple[str, Any]]) -> Optional[str]:
    """Canonical metric for a figure, taken from metric words on the figure's own line.

    Matching is confined to the line (so a metric in the next sentence/heading can't leak in),
    and picks the metric nearest *before* the number - the way "Revenue was $5M and EBITDA was
    $2M" reads. Ties prefer the longer surface, so "revenue growth" wins over "revenue" and
    "gross margin" over "margin" - distinct metrics that merely share a word stay separate.
    """
    line_start = text.rfind("\n", 0, fig_start) + 1
    line_end = text.find("\n", fig_end)
    if line_end == -1:
        line_end = len(text)

    before = text[line_start:fig_start]
    best = None  # (start_index, surface_length, canonical) - nearest-preceding wins on tuple cmp
    for canon, rx in aliases:
        for mm in rx.finditer(before):
            cand = (mm.start(), mm.end() - mm.start(), canon)
            if best is None or cand[:2] > best[:2]:
                best = cand
    if best is not None:
        return best[2]

    # No metric precedes the number on this line; accept the nearest one that follows it
    # (handles "$5M in revenue"), still confined to the same line.
    after = text[fig_end:line_end]
    best = None  # (-start_index, surface_length, canonical) - smallest start (nearest) wins
    for canon, rx in aliases:
        mm = rx.search(after)
        if mm:
            cand = (-mm.start(), mm.end() - mm.start(), canon)
            if best is None or cand[:2] > best[:2]:
                best = cand
    return best[2] if best is not None else None


def _extract_labeled_figures(text: str,
                             aliases: Optional[List[Tuple[str, Any]]] = None) -> List[Dict]:
    figures = []
    for m in _FIGURE_RE.finditer(text):
        num = m.group("n1") or m.group("n2")
        if not num:
            continue
        val = _parse_amount(num)
        if val is None:
            continue
        mag = (m.group("m1") or m.group("m2") or "").lower()
        if mag:
            val *= _MAGNITUDES.get(mag, 1)
        # Prefer a canonical metric key (order-independent) from the figure's surrounding
        # sentence; fall back to the positional last-few-words label when no metric is named.
        label = ""
        if aliases:
            label = _metric_key(text, m.start(), m.end(), aliases) or ""
        if not label:
            start = max(0, m.start() - 45)
            label = re.sub(r"\s+", " ", text[start:m.start()]).strip().lower()
            label = " ".join(_significant_words(label)[-3:])
        figures.append({"value": val, "raw": m.group(0).strip(), "label": label})
    return figures


# Numeric comparison tolerance: treat figures within 1% as the same value, so rounding
# ($4.9M vs $4.94M) and magnitude restatements ($5M vs $5,000,000) don't read as conflicts.
_FIGURE_TOL = 0.01


def _values_conflict(values, tol: float = _FIGURE_TOL) -> bool:
    vals = sorted(values)
    if len(vals) < 2 or vals[-1] == vals[0]:
        return False
    denom = max(abs(vals[-1]), abs(vals[0]), 1.0)
    return (vals[-1] - vals[0]) / denom > tol


def _table_body_conflict(tvals, bvals, tol: float = _FIGURE_TOL) -> bool:
    # A mismatch only when no table value is within tolerance of any narrative value.
    for t in tvals:
        for b in bvals:
            if abs(t - b) / max(abs(t), abs(b), 1.0) <= tol:
                return False
    return True


def _phase2_correctness(text: str, context: str, pat: Dict[str, Any]) -> List[Dict]:
    findings: List[Dict] = []
    n = 0

    # Mixed direction-of-change
    pos = pat.get("direction_positive", [])
    neg = pat.get("direction_negative", [])
    direction_issues = []
    for i, line in enumerate(text.split("\n")):
        low = line.lower()
        if any(re.search(p, low) for p in pos) and any(re.search(p, low) for p in neg):
            direction_issues.append(f"Line {i + 1}: {line.strip()[:120]}")
    if direction_issues:
        n += 1
        findings.append(_finding(
            f"COR-{n:03d}", "correctness", "mixed_direction_signals", MEDIUM,
            f"{len(direction_issues)} line(s) contain contradictory direction-of-change language",
            "These lines mix positive and negative direction words in one sentence, which may "
            "confuse the reader or signal an analytical slip.",
            evidence="\n".join(f"  - {d}" for d in direction_issues[:10]),
            remediation="Review each line; make the directional language match the underlying data.",
        ))

    aliases = _build_metric_aliases(pat)

    # Inconsistent labeled figures
    figs = _extract_labeled_figures(text, aliases)
    by_label: Dict[str, set] = {}
    for f in figs:
        if f["label"]:
            by_label.setdefault(f["label"], set()).add(round(f["value"], 2))
    inconsistent = [f"'{lab}': {sorted(vals)}" for lab, vals in by_label.items()
                    if _values_conflict(vals)]
    if inconsistent:
        n += 1
        findings.append(_finding(
            f"COR-{n:03d}", "correctness", "inconsistent_figures", MEDIUM,
            f"{len(inconsistent)} label(s) carry more than one value",
            "The same labeled quantity appears with different values in different places. One of "
            "them may be stale or wrong.",
            evidence="\n".join(f"  - {c}" for c in inconsistent[:10]),
            remediation="Reconcile each label to a single authoritative value (or distinguish the "
            "contexts explicitly).",
        ))

    # Extreme percentages (suppressed inside a growth/CAGR context, where >100% is normal:
    # "grew 250% year over year" is not an error, but a 250% margin is).
    growth_ctx = re.compile(pat.get("growth_context", r"(?!x)x"), re.IGNORECASE)
    extreme = []
    for m in re.finditer(r"(\d+(?:\.\d+)?)\s*%", text):
        try:
            val = float(m.group(1))
        except ValueError:
            continue
        if val <= 100:
            continue
        window = text[max(0, m.start() - 60):m.end() + 20]
        if growth_ctx.search(window):
            continue
        ctx = re.sub(r"\s+", " ", text[max(0, m.start() - 40):m.end() + 20]).strip()
        extreme.append(f"{m.group(1)}% in: ...{ctx}...")
    if extreme:
        n += 1
        findings.append(_finding(
            f"COR-{n:03d}", "correctness", "extreme_percentages", MEDIUM,
            f"{len(extreme)} percentage value(s) exceed 100% outside a growth context",
            "Unusually high percentages should be verified; some are legitimate (growth, "
            "coverage ratios) but others indicate a calculation error.",
            evidence="\n".join(f"  - {e}" for e in extreme[:10]),
            remediation="Verify each value against its source data.",
        ))

    # Hallucination - false access / capability claims
    hits = []
    for pattern, label in pat.get("hallucination_patterns", []):
        for m in re.finditer(pattern, text, re.IGNORECASE):
            ctx = re.sub(r"\s+", " ", text[max(0, m.start() - 25):m.end() + 45]).strip()
            hits.append(f'"{label}": ...{ctx}...')
    scope_pattern = pat.get("scope_none_pattern")
    scope_says_none = bool(scope_pattern and re.search(
        scope_pattern, (text + " " + context), re.IGNORECASE))
    if hits:
        sev = CRITICAL if scope_says_none else HIGH
        n += 1
        findings.append(_finding(
            f"COR-{n:03d}", "correctness", "hallucination", sev,
            f"{len(hits)} potential hallucination(s) - false access/interaction/testing claims",
            "The deliverable claims interactions, visits, surveys, or testing that may not have "
            "occurred. False claims of access are the most serious credibility risk in "
            "AI-assisted work - they misrepresent the evidence basis."
            + (" The stated scope says no such access occurred, which makes these claims a direct "
               "contradiction." if scope_says_none else ""),
            evidence="\n".join(f"  - {h}" for h in hits[:10]),
            remediation="Remove or correct every false access claim. Replace with an accurate "
            "description of the real basis (e.g. 'based on document review' or 'estimated from "
            "published benchmarks'). Verify each surviving claim against actual records.",
        ))

    # Table vs body figure mismatch
    table_lines, body_lines = [], []
    for line in text.split("\n"):
        (table_lines if line.strip().startswith("|") else body_lines).append(line)
    tmap: Dict[str, set] = {}
    bmap: Dict[str, set] = {}
    for f in _extract_labeled_figures("\n".join(table_lines), aliases):
        if f["label"]:
            tmap.setdefault(f["label"], set()).add(round(f["value"], 2))
    for f in _extract_labeled_figures("\n".join(body_lines), aliases):
        if f["label"]:
            bmap.setdefault(f["label"], set()).add(round(f["value"], 2))
    mismatches = [f"'{lab}': table {sorted(tmap[lab])} vs narrative {sorted(bmap[lab])}"
                  for lab in set(tmap) & set(bmap)
                  if _table_body_conflict(tmap[lab], bmap[lab])]
    if mismatches:
        n += 1
        findings.append(_finding(
            f"COR-{n:03d}", "correctness", "table_body_mismatch", HIGH,
            f"{len(mismatches)} labeled figure(s) appear to differ between a table and the narrative",
            "A figure shown in a table does not match the same labeled figure in the prose. This "
            "is a heuristic match: labels are inferred from the words near each number, so confirm "
            "the pairing before acting. Where the pairing is correct, one value is stale or wrong.",
            evidence="\n".join(f"  - {m}" for m in mismatches[:10]),
            remediation="Confirm the table and narrative refer to the same quantity, then "
            "reconcile them to a single authoritative value.",
        ))

    # Post-cutoff data (only when a cut-off year is derivable from the scope/context)
    ctx_years = [int(y) for y in re.findall(r"\b(?:19|20)\d{2}\b", context)]
    if ctx_years:
        cutoff = max(ctx_years)
        late = sorted({int(y) for y in re.findall(r"\b(?:19|20)\d{2}\b", text) if int(y) > cutoff})
        if late:
            n += 1
            findings.append(_finding(
                f"COR-{n:03d}", "correctness", "post_cutoff_data", MEDIUM,
                f"{len(late)} year(s) in the deliverable postdate the stated cut-off ({cutoff})",
                "The stated scope implies a data cut-off, but the body references later years. "
                "Confirm these are clearly-labeled projections, not actuals presented as fact.",
                evidence="  Years after cut-off: " + ", ".join(str(y) for y in late),
                remediation="Label post-cut-off references as forecasts, or remove data that "
                "could not have been available within the stated scope.",
            ))

    # Mixed currency (suppressed when an explicit conversion/FX basis is stated)
    currencies = _currencies_in(text)
    conversion_note = re.search(
        r"exchange\s+rate|conversion\s+(?:rate|basis)|converted\s+(?:at|to|into|using)|"
        r"constant\s+currency|reporting\s+currency|fx\s+rate|at\s+\d[\d.,]*\s*(?:per|/)\s*[A-Za-z$£€¥]",
        text, re.IGNORECASE)
    if len(currencies) >= 2 and not conversion_note:
        n += 1
        findings.append(_finding(
            f"COR-{n:03d}", "correctness", "mixed_currency", MEDIUM,
            f"{len(currencies)} currencies appear with no stated conversion basis",
            "Multiple currencies are used and no conversion basis, rate, or reporting currency was "
            "found. Totals and comparisons across them may be misleading.",
            evidence="  Currencies: " + ", ".join(currencies),
            remediation="State the reporting currency and the conversion basis/rate, or convert "
            "all figures to one currency.",
        ))

    # Mixed measurement basis: flag only when conflicting bases co-occur in one paragraph,
    # which is a stronger apples-to-oranges signal than scattered document-wide mentions.
    groups = pat.get("baseline_groups") or {}
    mixed_bases = []
    for group_name, alts in groups.items():
        if len(alts) < 2:
            continue
        for para in re.split(r"\n\s*\n", text):
            if sum(1 for alt in alts if re.search(alt, para, re.IGNORECASE)) >= 2:
                mixed_bases.append(group_name)
                break
    if mixed_bases:
        n += 1
        findings.append(_finding(
            f"COR-{n:03d}", "correctness", "mixed_baseline", LOW,
            f"{len(mixed_bases)} measurement basis/bases mixed within a single passage",
            "A passage mixes measurement bases (e.g. CAGR vs year-over-year, nominal vs real, "
            "annual vs monthly). This is a heuristic and is sometimes legitimate, but it is a "
            "frequent source of apples-to-oranges comparison.",
            evidence="  Mixed within a passage: " + "; ".join(mixed_bases),
            remediation="State which basis applies to each figure, or normalize to a single basis.",
        ))

    return findings


# ---------------------------------------------------------------------------
# Phase 2 - Excel model checks (formula-level; only for .xlsx)
# ---------------------------------------------------------------------------

_XL_ERROR_VALUES = ("#REF!", "#DIV/0!", "#VALUE!", "#NAME?", "#N/A", "#NULL!", "#NUM!")


def _phase2_excel(deliverable_path: str, advisory: bool = False) -> List[Dict]:
    """Formula-level checks that only make sense on a live .xlsx model."""
    if Path(deliverable_path).suffix.lower() != ".xlsx":
        return []
    try:
        sheets = read_xlsx_cells(deliverable_path)
    except Exception:  # noqa: BLE001 - never let a model-audit pass break the scan
        return []

    findings: List[Dict] = []
    n = 0

    # Formula error cells (always on - these are unambiguous broken calculations).
    errors = []
    for sh in sheets:
        for c in sh["cells"]:
            if c["error"]:
                errors.append(f"{sh['name']}!{c['ref']} = {c['error']}")
    if errors:
        n += 1
        findings.append(_finding(
            f"COR-XL-{n:03d}", "correctness", "formula_errors", HIGH,
            f"{len(errors)} cell(s) evaluate to a formula error",
            "Cells in the workbook resolve to spreadsheet errors (#REF!, #DIV/0!, etc.). A "
            "formula error usually means a broken reference or a divide-by-zero that silently "
            "corrupts every total downstream.",
            evidence="\n".join(f"  - {e}" for e in errors[:15]),
            remediation="Trace and fix each errored formula; confirm the dependent totals.",
        ))

    # Hardcoded constants sitting inside an otherwise-formula row (advisory: legitimate input
    # cells also look like this, so it is a heuristic). Off by default.
    if advisory:
        hardcoded = []
        for sh in sheets:
            rows: Dict[int, List[Dict]] = {}
            for c in sh["cells"]:
                rows.setdefault(c["row"], []).append(c)
            for cells in rows.values():
                fcols = [c["col"] for c in cells if c["has_formula"]]
                if len(fcols) < 2:
                    continue
                lo, hi = min(fcols), max(fcols)
                for c in cells:
                    if c["is_number"] and lo < c["col"] < hi:
                        hardcoded.append(f"{sh['name']}!{c['ref']}")
        if hardcoded:
            n += 1
            findings.append(_finding(
                f"COR-XL-{n:03d}", "correctness", "hardcoded_in_formula_range", MEDIUM,
                f"{len(hardcoded)} hardcoded number(s) inside a formula row",
                "These cells hold a typed-in number while the cells around them in the same row "
                "are formulas. That is a classic spot for an override that breaks when the model "
                "is rolled forward. This is a heuristic: genuine input cells look the same.",
                evidence="\n".join(f"  - {h}" for h in hardcoded[:15]),
                remediation="Confirm each is an intended input, not a formula that was typed "
                "over; convert overrides back to formulas where appropriate.",
            ))

    return findings


# ---------------------------------------------------------------------------
# Phase 3 - Adversarial
# ---------------------------------------------------------------------------

def _phase3_adversarial(text: str, pat: Dict[str, Any], advisory: bool = False) -> List[Dict]:
    findings: List[Dict] = []
    n = 0

    citation_hint = re.compile(pat.get("citation_hint", r"(?!x)x"), re.IGNORECASE)

    # Orphan assertions (advisory: high false-positive rate on real finance prose, where most
    # quantified claims are sourced elsewhere or in an accompanying model). Off by default.
    if advisory:
        orphans = []
        for para in re.split(r"\n\s*\n", text):
            s = para.strip()
            if len(s) < 60 or s.startswith("|") or s.startswith("#"):
                continue
            has_numbers = bool(re.search(
                rf"{_SYM_RE}\s*[\d,]+|\d+(?:\.\d+)?\s*%|\b\d{{1,3}}(?:,\d{{3}})+\b", s))
            if has_numbers and not citation_hint.search(s):
                orphans.append(s.split("\n")[0][:120])
        if orphans:
            sev = HIGH if len(orphans) > 10 else MEDIUM
            n += 1
            findings.append(_finding(
                f"ADV-{n:03d}", "adversarial", "orphan_assertions", sev,
                f"{len(orphans)} paragraph(s) state figures without a nearby source",
                "A hostile reader will challenge any quantified claim lacking a traceable source. "
                "These paragraphs assert numbers but cite nothing.",
                evidence="\n".join(f"  - {o}" for o in orphans[:10]),
                remediation="Add a source citation to each quantified claim (document, dataset, "
                "study, or stated assumption).",
            ))

        # Hedging density (advisory: finance writing legitimately hedges in almost every
        # sentence, so density alone is a weak signal). Off by default.
        hedges = []
        for pattern in pat.get("hedges", []):
            for m in re.finditer(pattern, text, re.IGNORECASE):
                ctx = re.sub(r"\s+", " ", text[max(0, m.start() - 25):m.end() + 25]).strip()
                hedges.append(f'"{m.group()}": ...{ctx}...')
        words = max(len(text.split()), 1)
        density = len(hedges) / words * 1000
        if density > 3.0:
            sev = HIGH if density > 6.0 else MEDIUM
            n += 1
            findings.append(_finding(
                f"ADV-{n:03d}", "adversarial", "hedging_language", sev,
                f"{len(hedges)} hedging word(s) ({density:.1f} per 1000 words)",
                "Heavy hedging signals uncertainty and invites challenge. Each hedge is a point "
                "that could be made precise.",
                evidence="\n".join(f"  - {h}" for h in hedges[:10]),
                remediation="Replace hedges with precise values or explicit ranges where the data "
                "permits; keep a hedge only where genuine uncertainty exists.",
            ))

    # Deferred analysis
    deferred = []
    for pattern, label in pat.get("deferred", []):
        for m in re.finditer(pattern, text, re.IGNORECASE):
            ctx = re.sub(r"\s+", " ", text[max(0, m.start() - 25):m.end() + 45]).strip()
            deferred.append(f'"{label}": ...{ctx}...')
    if deferred:
        sev = HIGH if len(deferred) > 5 else MEDIUM
        n += 1
        findings.append(_finding(
            f"ADV-{n:03d}", "adversarial", "deferred_analysis", sev,
            f"{len(deferred)} instance(s) of deferred or incomplete analysis",
            "These phrases mark places where the work stopped short. A reviewer asks: if the data "
            "was accessible, why was it not analyzed? If not, what was done to obtain it?",
            evidence="\n".join(f"  - {d}" for d in deferred[:10]),
            remediation="Complete the analysis now where the data is accessible; otherwise convert "
            "the gap into a specific, owned request with a deadline.",
        ))

    # Unqualified superlatives
    supers = []
    sup_re = re.compile(pat.get("superlative_pattern", r"(?!x)x"), re.IGNORECASE)
    qual_re = re.compile(pat.get("superlative_qualifier_pattern", r"(?!x)x"), re.IGNORECASE)
    for m in sup_re.finditer(text):
        window = text[max(0, m.start() - 100):m.end() + 100]
        if not qual_re.search(window):
            ctx = re.sub(r"\s+", " ", text[max(0, m.start() - 25):m.end() + 35]).strip()
            supers.append(f'"{m.group()}": ...{ctx}...')
    if supers:
        sev = MEDIUM if len(supers) > 5 else LOW
        n += 1
        findings.append(_finding(
            f"ADV-{n:03d}", "adversarial", "unqualified_superlatives", sev,
            f"{len(supers)} superlative claim(s) without a benchmark",
            "Superlatives without a comparator invite the question 'compared to what?'",
            evidence="\n".join(f"  - {s}" for s in supers[:10]),
            remediation="Qualify each superlative with a time period, peer set, or benchmark.",
        ))

    # Confidentiality / PII exposure (advisory: pattern matches misfire on extracted text,
    # especially from PDFs with imperfect font decoding). Off by default. Only actual personal
    # data or secrets count as an exposure; "confidential" markers are reported as context.
    if advisory:
        pii_hits = []
        marker_hits = []
        worst = LOW
        for pattern, label, sev in _PII_PATTERNS:
            for m in re.finditer(pattern, text):
                snippet = re.sub(r"\s+", " ", m.group(0)).strip()[:60]
                if label == "confidentiality marker":
                    marker_hits.append(f"{label}: {snippet}")
                else:
                    pii_hits.append(f"{label}: {snippet}")
                    if _SEVERITY_ORDER[sev] < _SEVERITY_ORDER[worst]:
                        worst = sev
        if pii_hits:
            n += 1
            evidence = [f"  - {h}" for h in pii_hits[:10]]
            if marker_hits:
                evidence.append(f"  (also present: {len(marker_hits)} confidentiality marker(s))")
            findings.append(_finding(
                f"ADV-{n:03d}", "adversarial", "confidentiality_pii", worst,
                f"{len(pii_hits)} possible personal-data / secret exposure(s)",
                "The deliverable contains personal data, credentials, or keys. These are a leak "
                "and compliance risk if the document is shared."
                + (" The document is also marked confidential, so handle accordingly."
                   if marker_hits else ""),
                evidence="\n".join(evidence),
                remediation="Redact or remove personal data and secrets before distribution.",
            ))

    return findings


# ---------------------------------------------------------------------------
# Assemble + score
# ---------------------------------------------------------------------------

# Scoring is severity-weighted, not count-driven. Blocking findings (Critical/High) deduct
# in full; advisory findings (Medium/Low) are capped so that a long tail of minor, heuristic
# items cannot collapse the score of an otherwise solid document. The headline signal is the
# verdict band, which is driven by the worst open severity, not by the raw percentage.
ADVISORY_CAP_PHASE = 24
ADVISORY_CAP_OVERALL = 36

VERDICT_LABELS = {
    "not_ready": "Not ready - blocking issues to resolve",
    "needs_work": "Needs work - significant items to address",
    "review_minor": "Solid - minor, fixable items remain",
    "clean": "Clean - no automated findings",
}


def active_findings(findings: List[Dict]) -> List[Dict]:
    """Findings that still count: everything except those explicitly dismissed."""
    return [f for f in findings if f.get("status", "open") != "dismissed"]


def score_findings(findings: List[Dict], advisory_cap: int = ADVISORY_CAP_OVERALL) -> int:
    """Severity-weighted score: blocking deducts fully, advisory deducts up to a cap."""
    fs = active_findings(findings)
    blocking = sum(_PENALTY.get(f.get("severity"), 0)
                   for f in fs if f.get("severity") in (CRITICAL, HIGH))
    advisory = min(advisory_cap,
                   sum(_PENALTY.get(f.get("severity"), 0)
                       for f in fs if f.get("severity") in (MEDIUM, LOW)))
    return max(0, 100 - blocking - advisory)


def verdict_for(findings: List[Dict]) -> str:
    """Readiness band driven by the worst open severity, not by the raw score."""
    fs = active_findings(findings)
    if any(f.get("severity") == CRITICAL for f in fs):
        return "not_ready"
    if any(f.get("severity") == HIGH for f in fs):
        return "needs_work"
    if fs:
        return "review_minor"
    return "clean"


TOOL_VERSION = "0.1.1"

# The reasoning-only sub-checks the agent must complete in Step 2. The scanner cannot perform
# these, so it seeds them as "pending"; the agent updates each to "completed"/"not_applicable".
AGENT_CHECK_KEYS = (
    "figure_re_verification", "internal_recompute", "challenge_simulation",
    "cross_document", "single_source_risk", "hallucination",
)

def load_checklist(path: Optional[str]) -> Tuple[Optional[Dict[str, Any]], Optional[str]]:
    """Return (checklist, warning) for a checklist JSON file.

    Checklists are generated per deliverable (usually by the agent in Step 1b) rather than
    shipped, so this loads a path to a JSON file of the shape
    {"name": ..., "sections": [{"title", "severity", "keywords": [...]}, ...]}.
    Loading is fail-soft: a bad or missing file warns and skips checklist coverage.
    """
    if not path:
        return None, None
    src = Path(path)
    if not src.is_file():
        return None, f"Checklist not found: '{path}'. Skipping checklist coverage."
    try:
        data = json.loads(src.read_text(encoding="utf-8"))
        if not isinstance(data, dict) or not isinstance(data.get("sections"), list):
            raise ValueError("checklist must be an object with a 'sections' list")
        return data, None
    except Exception as exc:  # noqa: BLE001 - surface a clean warning, never hard-fail
        return None, f"Could not read checklist '{src}': {exc}. Skipping checklist coverage."


def _checklist_findings(text: str, checklist: Dict[str, Any]) -> List[Dict]:
    """Flag expected sections (per a domain checklist) whose keywords never appear."""
    findings: List[Dict] = []
    low = text.lower()
    name = checklist.get("name", "checklist")
    n = 0
    for sec in checklist.get("sections", []):
        kws = [str(k).lower() for k in sec.get("keywords", []) if k]
        if not kws or any(k in low for k in kws):
            continue
        sev = sec.get("severity", MEDIUM)
        if sev not in (CRITICAL, HIGH, MEDIUM, LOW):
            sev = MEDIUM
        title = sec.get("title", "(untitled)")
        n += 1
        findings.append(_finding(
            f"COV-CL-{n:03d}", "coverage", "missing_expected_section", sev,
            f"Expected section appears to be missing: {title}",
            f"The '{name}' checklist expects coverage of '{title}', but none of its keywords "
            "appear in the deliverable. Confirm the topic was addressed (perhaps under another "
            "heading) or is genuinely out of scope.",
            evidence="  Looked for any of: " + ", ".join(kws[:8]),
            remediation=f"Add a section covering {title}, or state explicitly why it is out of scope.",
        ))
    return findings


def _category_counts(register: Dict[str, Any]) -> Dict[str, int]:
    counts: Dict[str, int] = {}
    for ph in register.get("phases", {}).values():
        for f in ph.get("findings", []):
            counts[f.get("category", "?")] = counts.get(f.get("category", "?"), 0) + 1
    return counts


def diff_registers(old: Dict[str, Any], new: Dict[str, Any]) -> Dict[str, Any]:
    """Category-level diff between a baseline register and the current one (robust to id churn)."""
    oc, nc = _category_counts(old), _category_counts(new)
    resolved = sorted(c for c in oc if c not in nc)
    introduced = sorted(c for c in nc if c not in oc)
    changed = {c: (oc[c], nc[c]) for c in sorted(set(oc) & set(nc)) if oc[c] != nc[c]}
    return {
        "resolved": resolved,
        "new": introduced,
        "changed": changed,
        "old_verdict": old.get("summary", {}).get("verdict"),
        "new_verdict": new.get("summary", {}).get("verdict"),
        "old_score": old.get("summary", {}).get("overall_confidence"),
        "new_score": new.get("summary", {}).get("overall_confidence"),
    }


_FORMAT_LABELS = {
    ".docx": "Word (.docx)", ".xlsx": "Excel (.xlsx)", ".pptx": "PowerPoint (.pptx)",
    ".pdf": "PDF (.pdf)",
    ".md": "Markdown", ".markdown": "Markdown", ".txt": "Plain text",
    ".rst": "reStructuredText", ".text": "Plain text",
}


def _format_label(path: str) -> str:
    return _FORMAT_LABELS.get(Path(path).suffix.lower(), "Plain text")


def build_register(deliverable_path: str, sources_dir: Optional[str], context: str,
                   pat: Optional[Dict[str, Any]] = None,
                   advisory_checks: bool = False,
                   checklist: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
    if pat is None:
        pat, _ = load_patterns()
    text = load_deliverable_text(deliverable_path)
    extraction_ok, extraction_note = extraction_quality(text)
    scan_text = _strip_noise(text)
    p1 = _phase1_coverage(text, sources_dir)
    if checklist:
        p1 = p1 + _checklist_findings(text, checklist)
    p2 = _phase2_correctness(scan_text, context or "", pat)
    p2 = p2 + _phase2_excel(deliverable_path, advisory_checks)
    p3 = _phase3_adversarial(scan_text, pat, advisory_checks)
    allf = sorted(p1 + p2 + p3, key=lambda f: _SEVERITY_ORDER.get(f["severity"], 99))
    active = active_findings(allf)
    summary = {
        "verdict": verdict_for(allf),
        "verdict_label": VERDICT_LABELS[verdict_for(allf)],
        "total_findings": len(allf),
        "blocking": sum(1 for f in active if f["severity"] in (CRITICAL, HIGH)),
        "advisory": sum(1 for f in active if f["severity"] in (MEDIUM, LOW)),
        "critical": sum(1 for f in allf if f["severity"] == CRITICAL),
        "high": sum(1 for f in allf if f["severity"] == HIGH),
        "medium": sum(1 for f in allf if f["severity"] == MEDIUM),
        "low": sum(1 for f in allf if f["severity"] == LOW),
        "auto_fixable": sum(1 for f in allf if f.get("auto_fixable")),
        "coverage_score": score_findings(p1, ADVISORY_CAP_PHASE),
        "correctness_score": score_findings(p2, ADVISORY_CAP_PHASE),
        "adversarial_score": score_findings(p3, ADVISORY_CAP_PHASE),
        "overall_confidence": score_findings(allf, ADVISORY_CAP_OVERALL),
    }
    return {
        "anxiety_register_id": str(uuid.uuid4()),
        "deliverable_path": str(deliverable_path),
        "deliverable_format": _format_label(deliverable_path),
        "language": pat.get("language", "en"),
        "context": context or "",
        "extraction_ok": extraction_ok,
        "extraction_warning": extraction_note,
        "timestamp": datetime.now(timezone.utc).isoformat(),
        "tool": {
            "name": "anxiety",
            "version": TOOL_VERSION,
            "advisory_checks": bool(advisory_checks),
            "checklist": (checklist.get("name") if checklist else None),
        },
        "phases": {
            "coverage": {"score": summary["coverage_score"], "findings": p1},
            "correctness": {"score": summary["correctness_score"], "findings": p2},
            "adversarial": {"score": summary["adversarial_score"], "findings": p3},
        },
        "agent_checks": {k: {"status": "pending", "items_checked": 0, "note": ""}
                         for k in AGENT_CHECK_KEYS},
        "summary": summary,
    }


def write_markdown(register: Dict, path: str) -> None:
    s = register["summary"]
    L: List[str] = []
    L += ["# Anxiety Register", "",
          f"**Deliverable:** {register['deliverable_path']}",
          f"**Format:** {register.get('deliverable_format', 'Plain text')}",
          f"**Language pack:** {register.get('language', 'en')}",
          f"**Generated:** {register['timestamp']}",
          f"**Register ID:** {register['anxiety_register_id']}"]
    if register.get("context"):
        L.append(f"**Stated scope:** {register['context']}")
    L += ["", "---", "", "## Summary", "",
          f"**Verdict: {s.get('verdict_label', s.get('verdict', ''))}**", "",
          "| Metric | Value |", "|---|---|",
          f"| Verdict | **{s.get('verdict_label', '')}** |",
          f"| Blocking findings (Critical/High) | {s.get('blocking', s['critical'] + s['high'])} |",
          f"| Advisory findings (Medium/Low) | {s.get('advisory', s['medium'] + s['low'])} |",
          f"| Severity-weighted score | {s['overall_confidence']}/100 |",
          f"| Coverage Score | {s['coverage_score']}/100 |",
          f"| Correctness Score | {s['correctness_score']}/100 |",
          f"| Adversarial Score | {s['adversarial_score']}/100 |",
          f"| Total Findings | {s['total_findings']} |",
          f"| Critical | {s['critical']} |",
          f"| High | {s['high']} |",
          f"| Medium | {s['medium']} |",
          f"| Low | {s['low']} |", "",
          "_The verdict is the headline: it reflects the worst open severity. The "
          "severity-weighted score caps the contribution of minor, heuristic findings so a long "
          "tail of low-severity items doesn't sink an otherwise sound document._", ""]
    for key, label in [
        ("coverage", "Phase 1 - Coverage: What Did I Miss?"),
        ("correctness", "Phase 2 - Correctness: Is It Actually Right?"),
        ("adversarial", "Phase 3 - Adversarial: What Would a Hostile Reader Attack?"),
    ]:
        ph = register["phases"][key]
        L += ["---", "", f"## {label}", "", f"**Phase Score: {ph['score']}/100**", ""]
        if not ph["findings"]:
            L += ["No findings. All checks passed.", ""]
            continue
        for f in ph["findings"]:
            L += [f"### [{f['severity'].upper()}] {f['id']}: {f['title']}", "", f["detail"], ""]
            if f.get("evidence"):
                L += ["**Evidence:**", "", f["evidence"], ""]
            if f.get("remediation"):
                L += [f"**Remediation:** {f['remediation']}", ""]
    Path(path).write_text("\n".join(L), encoding="utf-8")


def main() -> int:
    ap = argparse.ArgumentParser(description="Anxiety scan for any knowledge-work deliverable.")
    ap.add_argument("deliverable", help="Path to the deliverable (.docx/.xlsx/.pptx/.pdf/.md/.txt).")
    ap.add_argument("--sources", help="Directory of source files that should have informed it.")
    ap.add_argument("--context", default="", help="Scope/context note (e.g. 'desk research only').")
    ap.add_argument("--lang", default="en", help="Language pack code (default: en).")
    ap.add_argument("--patterns", help="Path to a custom pattern pack JSON (overrides --lang).")
    ap.add_argument("--advisory-checks", action="store_true",
                    help="Also run the high-false-positive heuristic checks (orphan assertions, "
                         "hedging density, PII/secret patterns, hardcoded model cells). Off by default.")
    ap.add_argument("--checklist", help="Path to a coverage checklist JSON (usually one the "
                                        "agent generated for this deliverable; see SKILL.md "
                                        "Step 1b).")
    ap.add_argument("--baseline", help="A prior <name>.anxiety.json to diff this run against.")
    ap.add_argument("--out", help="Output directory (defaults next to the deliverable).")
    args = ap.parse_args()

    print("\U0001F630\U0001F631 aaahhh I have anxiety!! \U0001F630\U0001F4A6  "
          "...let me re-interrogate this work.\n")

    if not os.path.isfile(args.deliverable):
        print(f"Error: deliverable not found: {args.deliverable}")
        return 2

    pat, warning = load_patterns(args.lang, args.patterns)
    checklist, checklist_warning = load_checklist(args.checklist)

    out_dir = args.out or str(Path(args.deliverable).parent)
    os.makedirs(out_dir, exist_ok=True)
    stem = Path(args.deliverable).stem
    json_path = os.path.join(out_dir, f"{stem}.anxiety.json")
    md_path = os.path.join(out_dir, f"{stem}.anxiety.md")

    print(f"Running anxiety scan on: {args.deliverable}")
    print(f"  Format:   {_format_label(args.deliverable)}")
    print(f"  Language: {pat.get('language', 'en')}")
    if args.sources:
        print(f"  Sources:  {args.sources}")
    if args.context:
        print(f"  Context:  {args.context}")
    if checklist:
        print(f"  Checklist: {checklist.get('name', args.checklist)}")
    if warning:
        print(f"  Note:     {warning}")
    if checklist_warning:
        print(f"  Note:     {checklist_warning}")
    print()

    try:
        register = build_register(args.deliverable, args.sources, args.context, pat,
                                  advisory_checks=args.advisory_checks, checklist=checklist)
    except ExtractionError as exc:
        print(f"Error: {exc}")
        return 2
    Path(json_path).write_text(json.dumps(register, indent=2, ensure_ascii=False), encoding="utf-8")
    write_markdown(register, md_path)

    if args.baseline and os.path.isfile(args.baseline):
        try:
            old = json.loads(Path(args.baseline).read_text(encoding="utf-8"))
            d = diff_registers(old, register)
            diff_path = os.path.join(out_dir, f"{stem}.anxiety.diff.md")
            dl = [f"# Anxiety diff vs baseline", "",
                  f"Baseline: {args.baseline}", "",
                  f"- Verdict: {d['old_verdict']} -> {d['new_verdict']}",
                  f"- Score: {d['old_score']} -> {d['new_score']}",
                  f"- Resolved categories: {', '.join(d['resolved']) or 'none'}",
                  f"- New categories: {', '.join(d['new']) or 'none'}",
                  "- Changed counts: " + (", ".join(f"{c} {o}->{n}" for c, (o, n) in d['changed'].items()) or "none"),
                  ""]
            Path(diff_path).write_text("\n".join(dl), encoding="utf-8")
            print("Diff vs baseline:")
            print(f"  Verdict: {d['old_verdict']} -> {d['new_verdict']}")
            print(f"  Resolved: {', '.join(d['resolved']) or 'none'}")
            print(f"  New:      {', '.join(d['new']) or 'none'}")
            print(f"  Diff:     {diff_path}")
            print()
        except (ValueError, OSError) as exc:
            print(f"  Note:     could not diff baseline: {exc}\n")

    if not register.get("extraction_ok", True):
        print("  ⚠️  Extraction quality warning: " + register.get("extraction_warning", ""))
        print("     Findings below may be unreliable. Prefer a higher-fidelity reader "
              "(see SKILL.md) and re-run on its output.\n")

    s = register["summary"]
    print("Anxiety Register Results")
    print("=" * 50)
    print(f"  Verdict:             {s.get('verdict_label', s.get('verdict', ''))}")
    print(f"  Blocking / Advisory: {s.get('blocking', 0)} / {s.get('advisory', 0)}")
    print(f"  Weighted score:      {s['overall_confidence']}/100")
    print(f"  Coverage Score:      {s['coverage_score']}/100")
    print(f"  Correctness Score:   {s['correctness_score']}/100")
    print(f"  Adversarial Score:   {s['adversarial_score']}/100")
    print(f"  Total Findings:      {s['total_findings']}  "
          f"(C:{s['critical']} H:{s['high']} M:{s['medium']} L:{s['low']})")
    if not args.advisory_checks:
        print("  (advisory heuristic checks off; add --advisory-checks to include them)")
    print()
    print(f"  JSON:     {json_path}")
    print(f"  Markdown: {md_path}")
    return 0 if s["critical"] == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())
