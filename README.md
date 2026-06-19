<p align="center">
  <img src="assets/anxiety-mascot.svg" alt="Anxious cartoon mascot worrying about unsourced figures" width="560">
</p>

<h1 align="center">Anxiety Check</h1>

<p align="center">
  <em>What did I miss? Is this actually right? What would a hostile reader attack?</em>
</p>

<p align="center">
  Developed by <strong>Arsalan Pardesi</strong>
</p>

<p align="center">
  <img src="https://img.shields.io/badge/license-MIT-green.svg" alt="License: MIT">
  <img src="https://img.shields.io/badge/python-3.8%2B-blue.svg" alt="Python 3.8+">
  <img src="https://img.shields.io/badge/dependencies-stdlib%20only-success.svg" alt="Dependencies: stdlib only">
  <img src="https://img.shields.io/badge/type-agent%20skill-orange.svg" alt="Type: agent skill">
  <img src="https://img.shields.io/badge/mode-diagnostic%20only-lightgrey.svg" alt="Mode: diagnostic only">
</p>

> Re-interrogate any completed knowledge-work deliverable across three phases: **Coverage**, **Correctness**, and **Adversarial**. Catches hallucinated access claims, unsourced figures, deferred analysis, and thin coverage. Diagnostic only, never edits your work.

`anxiety` is an [agent skill] that gives an AI assistant a structured, repeatable methodology for stress-testing finished work the way an anxious expert would the night before delivery: *What did I miss? Is this actually right? What would a hostile reader attack?*

Most review tooling checks whether a document is *well-formed* (headings, formatting, length). The anxiety check interrogates whether the **underlying work is thorough, correct, and defensible**.

| Question | Phase |
|---|---|
| Were all available sources actually used? | Coverage |
| Do the stated figures and claims hold up? Did we invent anything? | Correctness |
| Would a hostile, expert reader be able to attack it? | Adversarial |

A formatting review asks *"is this written well?"* The anxiety check asks *"is this **true** and **complete**, and will it **survive challenge**?"*

---

## Why it exists

AI-assisted knowledge work has a specific failure mode: it produces output that *sounds* authoritative regardless of whether the evidence supports it. Reports claim interviews that never happened, cite sources that don't exist, attribute suspiciously precise statistics to named authorities, and quietly defer the hard analysis. These are exactly the things that fail under scrutiny, and exactly the things a polish-focused review misses.

The anxiety check is built to catch them.

## What it works on

Any completed knowledge-work deliverable: a consulting report, research write-up, legal memo, strategy doc, due-diligence report, technical design, grant proposal, investment memo, board paper, or academic paper.

**Supported file types:**

| Type | Extensions | How it is read |
|---|---|---|
| Word | `.docx` | Headings, paragraphs, and tables extracted to text |
| Excel | `.xlsx` | One section per sheet; rows read as table lines |
| PowerPoint | `.pptx` | One section per slide; slide title becomes a heading |
| PDF | `.pdf` | Text layer extracted from content streams (best-effort) |
| Markdown / text | `.md` `.markdown` `.txt` `.rst` | Read as-is |

Office files are parsed as ZIP + XML and PDF content streams are decoded via `zlib`, using only the Python standard library (`zipfile` + `xml.etree` + `zlib`), so there are **no third-party dependencies**. This makes the bundled extractor portable, but it's best-effort: the Office reader can miss complex layouts, charts, embedded objects, comments, or tracked changes, and the PDF reader (which undoes the Flate/LZW/ASCII85/ASCIIHex/RunLength filters) can't read scanned/image-only PDFs (no OCR), encrypted files, predictor-encoded or image streams, or custom/CID font encodings. Legacy binary formats (`.doc`, `.xls`, `.ppt`) aren't parsed directly.

So for **any** file type — not just PDF — the skill prefers a higher-fidelity reader when the harness has one, and only falls back to the bundled extractor when it doesn't:

- **Word / Excel / PowerPoint** — native file reading, `pandoc`, `markitdown`, LibreOffice, or `python-docx`/`openpyxl`/`python-pptx`.
- **PDF** — native file reading, `pdftotext`/poppler, `pandoc`, `markitdown`, `pdfminer`/`pypdf`, or OCR for scanned pages.

When a better reader is used, the skill saves its output to a `.md`/`.txt`, runs the scanner on that file, and records which extractor was used so the findings stay traceable. The bundled stdlib extractor is always there as the fallback; when even that can't read a file and no better tool exists, export to a modern format or to text/markdown first.

---

## How it works

The skill combines a fast, deterministic scanner with agent-mediated reasoning checks.

> **Where the skill lives:** the shippable skill (SKILL.md, `scripts/`, `references/`, `patterns/`) is in the nested [`anxiety/`](anxiety) folder. Run the commands below from inside it (`cd anxiety`), or prefix the paths with `anxiety/`.

```
deliverable ──▶ anxiety_scan.py ──▶ pattern findings (structure, hallucination patterns, ...)
                      │
                      ▼
              agent deep checks ──▶ figure re-verification, challenge simulation,
                      │              cross-doc reconciliation, single-source risk
                      ▼
            Anxiety Register (JSON + Markdown) ──▶ scored summary + remediation
```

### Step 1 — Portable scanner (pattern-based pass)

The bundled scanner is stdlib-only Python (no third-party deps, no database, no network), so it runs in **any harness with a shell**. It performs the structural and pattern checks for all three phases.

```bash
python scripts/anxiety_scan.py <deliverable.docx|.xlsx|.pptx|.pdf|.md|.txt> \
  [--sources <dir>] \
  [--context "<scope note, e.g. 'desk research only; no interviews conducted'>"] \
  [--lang <code>] \
  [--patterns <file.json>] \
  [--advisory-checks] \
  [--checklist <path.json>] \
  [--baseline <prior.anxiety.json>] \
  [--out <output_dir>]
```

Two optional flags help with coverage and re-runs: `--checklist` takes a JSON checklist and flags expected sections that never appear. Rather than ship a fixed set of domain checklists, the agent generates one tailored to the deliverable (its type, audience, and the decision it supports) and saves it as `<name>.anxiety.checklist.json` before the scan, so the bar fits *this* piece of work rather than a generic archetype. `--baseline` diffs the run against a prior register and writes a `<name>.anxiety.diff.md` so you can re-run after fixes and see what cleared.

The scanner auto-detects the file type and extracts text before analysing it. To inspect just the extracted text (useful for debugging):

```bash
python scripts/extract_text.py <file> [out.md]
```

**Outputs** (written next to the deliverable, or to `--out`):

- `<name>.anxiety.json` — structured register with every finding
- `<name>.anxiety.md` — human-readable summary

**What the scanner catches automatically:**

*Phase 1 — Coverage*
- **Unreferenced sources** — files in `--sources` whose subject never appears in the deliverable.
- **Thin sections** — headed sections with too little substantive content.
- **Missing expected sections** — with `--checklist`, expected topics whose keywords never appear.

*Phase 2 — Correctness*
- **Mixed direction-of-change signals** — one sentence asserting both up and down for a metric.
- **Extreme / implausible percentages** — values >100% outside a growth/CAGR context (a figure next to cues like "grew", "rose", "CAGR", or "year over year" is treated as normal growth, not an error).
- **Inconsistent figures** — the same metric stated with different values. Figures are keyed by the metric named nearest before them on the line (canonical names, longest match wins), so re-wordings of one metric are compared together while distinct metrics that share a word — `revenue` vs `revenue growth`, `gross` vs `operating margin` — stay separate.
- **Table / body mismatch** — a labeled figure that differs between a table and the prose.
- **Post-cutoff data** — years in the body that postdate the stated scope cut-off (with `--context`).
- **Mixed currency** — two or more currencies used without a stated conversion basis.
- **Mixed measurement basis** — CAGR vs YoY, nominal vs real, or annual vs monthly in one document.
- **Formula errors** (`.xlsx`) — cells that resolve to `#REF!`, `#DIV/0!`, and the like.
- **Hardcoded cells in a formula row** (`.xlsx`) — typed-in numbers among formulas. *(advisory)*
- **Hallucination patterns** — false claims of interviews, site visits, surveys, or testing.

Figure comparisons tolerate rounding: values within 1%, and magnitude restatements like `$5M` vs `$5,000,000`, are treated as the same number rather than a conflict.

After extraction, a best-effort heuristic checks the text isn't garbled (replacement characters, a high share of unusual symbols, or lost word spacing — common when a PDF uses an unsupported font/encoding). If it looks garbled the run sets `extraction_ok: false` in the register and prints a warning, so findings from a bad extraction aren't trusted silently. Digits and ordinary punctuation don't count against it, so numeric tables and spreadsheets aren't mistaken for garbage. It never blocks the scan.

*Phase 3 — Adversarial*
- **Deferred analysis** — "should be analyzed", "requires further review", "was not available".
- **Unqualified superlatives** — "the highest", "the best" with no benchmark or comparator.
- **Orphan assertions** — quantified claims with no nearby source or citation. *(advisory)*
- **Hedging-language density** — "broadly", "approximately", "largely", etc. *(advisory)*
- **Confidentiality / PII exposure** — emails, phone numbers, credentials, keys. *(advisory)*

The three *advisory* checks misfire often on real finance prose, so they're **off by default**; add `--advisory-checks` to include them. Everything else runs on every scan. And the scanner overall is an accelerant: on substantial analytical work, the findings that matter most come from the agent's reasoning in Step 2, not the regex layer.

Pattern-based checks run on a de-noised copy of the text: fenced code blocks, inline code spans, quoted spans (`"..."`), block quotes, and sections titled *example/sample/illustrative/template* are stripped first so quoted or illustrative text isn't mistaken for the author's own claim. This is best-effort, not exhaustive.

If no Python is available, the agent performs the equivalent checks manually using the same methodology. The scanner is an accelerant, not a hard dependency.

### Step 2 — Agent-mediated deep checks

The scanner catches structure and patterns; the agent adds the checks that require reasoning, each grounded in a verbatim quote plus the source it was checked against (anything unverifiable is recorded as low, not asserted):

- **Figure / claim re-verification** of the most load-bearing numbers (re-derive from source, flag discrepancies).
- **Internal arithmetic recomputation** — recompute the document's own math from the numbers it states (totals vs parts, EV/net-debt build-ups, margins and ratios, CAGR vs endpoints, consistent restatements). Needs no sources, so it works even when none are supplied — and it's the check IC/QoE work most wants.
- **Challenge simulation** of the top conclusions (write the strongest counter-argument, check whether the work pre-empts it).
- **Cross-document reconciliation** across related deliverables.
- **Anchoring / single-source risk** for material claims that rest on one unverified source.
- **Hallucination check** — false access claims, fabricated citations, invented data with false precision, phantom capabilities, and temporal inconsistencies.

For the numeric checks the agent is told to **use a real compute tool rather than mental math** — code execution, a Python/shell interpreter, `bc`, or the spreadsheet itself — and to quote the expression and inputs it ran. Mental arithmetic on a precision check both misses real errors and invents false ones.

The register also carries an `agent_checks` ledger that starts every reasoning check as `pending`; the agent marks each `completed` or `not_applicable` as it goes. Running `python scripts/validate_register.py <name>.anxiety.json --require-complete` fails while any check is still pending, so a clean-looking register can't be presented before the reasoning actually happened.

### Step 3 — Scoring & verdict

The headline is a **verdict band**, driven by the worst open finding — not a single percentage, because enough minor findings would otherwise sink any thorough review and punish completeness:

| Verdict | Trigger |
|---|---|
| **Not ready** | any open Critical |
| **Needs work** | any open High (no Critical) |
| **Solid — minor items** | only Medium/Low open |
| **Clean** | no findings |

A severity-weighted **score (0–100)** is still reported, but it's built not to collapse from volume: blocking findings (Critical/High) deduct in full (−25 / −15), while advisory findings (Medium/Low) deduct −8 / −3 only up to a cap (24 per phase, 36 overall). Dismissed findings don't count, so triaging false positives helps the score. Act on the verdict and the blocking count; read the number as a rough signal.

### Language packs

The language-specific cues the scanner looks for (hedges, deferred-analysis phrases, hallucination patterns, direction words, citation hints) live in JSON pattern packs under [`anxiety/patterns/`](anxiety/patterns). English ships built-in and as [`anxiety/patterns/en.json`](anxiety/patterns/en.json).

To run against a deliverable written in another language, copy `patterns/en.json`, translate the cue lists, save it as `patterns/<code>.json`, and pass `--lang <code>` (or point `--patterns` at any pack file):

```bash
python scripts/anxiety_scan.py rapport.docx --lang fr
```

If a pack is missing or unreadable the scanner falls back to the built-in English patterns and prints a note, so a wrong `--lang` never breaks a run. Structural checks (thin sections, inconsistent figures, extreme percentages) and the broadened currency parsing are language-independent and run regardless.

---

## Quick start

```bash
# Markdown / text
python scripts/anxiety_scan.py report.md

# Word, with the sources that should have informed it and a scope note
python scripts/anxiety_scan.py report.docx \
  --sources ./research_inputs \
  --context "desk research only; no interviews conducted"

# Excel, writing results to a chosen directory
python scripts/anxiety_scan.py model.xlsx --out ./review
```

### Example output

```
Anxiety Register Results
==================================================
  Verdict:             Not ready - blocking issues to resolve
  Blocking / Advisory: 3 / 4
  Weighted score:      18/100
  Coverage Score:      85/100
  Correctness Score:   67/100
  Adversarial Score:   66/100
  Total Findings:      7  (C:1 H:2 M:3 L:1)
  (advisory heuristic checks off; add --advisory-checks to include them)

  JSON:     report.anxiety.json
  Markdown: report.anxiety.md
```

The process exits `0` when no critical findings are present and `1` when at least one critical finding is found, so it can gate a CI step if desired.

---

## The Anxiety Register

Both the scanner and the agent write to the same JSON register, so automated and reasoning findings live together. Each finding has a stable shape:

```json
{
  "id": "COR-AGENT-001",
  "phase": "correctness",
  "category": "figure_re_verification",
  "severity": "high",
  "title": "Revenue figure discrepancy",
  "detail": "Executive summary states $4.9M; recomputed from source = $5.4M (10% gap).",
  "evidence": "source/financials.csv row 14 sums to 5,412,330",
  "remediation": "Correct the figure to $5.4M or reconcile the difference.",
  "auto_fixable": false,
  "status": "open"
}
```

Full schema and the category list are in [`anxiety/references/register-schema.md`](anxiety/references/register-schema.md).

---

## Diagnostic only — no automatic fixes

The anxiety check is a **diagnostic tool**, not a repair tool.

- The scanner never modifies the deliverable, the sources, or any user work product. It only writes its own `*.anxiety.json` / `*.anxiety.md` register files.
- The agent never applies fixes during the check unless the user **explicitly asks** to fix a specific item or category.
- After presenting results, it waits for instruction.

**Correct flow:** run check → present findings → user requests a fix → apply that fix → optionally re-run to show improvement.

**Prohibited flow:** run check → silently rewrite the deliverable.

## What it does *not* do

- It does not second-guess legitimate professional judgment calls.
- It does not rewrite the deliverable; it flags issues for you to decide on.
- It does not replace a domain expert's substantive review; it complements it with systematic coverage, correctness, and adversarial pressure.
- It is not part of any automated pipeline; it is user-initiated only.

---

## Scope & limitations (v0.1)

This is an early release. The scanner is honest about what it does and doesn't do:

- **Language.** The pattern-based checks are tuned for English and ship as a pack you can copy and translate (see [Language packs](#language-packs)). Run it on another language without a matching pack and the Correctness/Adversarial pattern checks will under-report; the structural checks still apply.
- **Numbers & currency.** Figure parsing now recognises common symbols (`£ $ € ¥ ₹ ...`), ISO codes (`USD`, `EUR`, `GBP`, ...), magnitude words (`k`/`m`/`bn`/`trillion`), and both US (`1,234.56`) and European (`1.234,56`, `1 234,56`) formats. Locale-specific conventions outside these (e.g. Indian `lakh`/`crore`) aren't yet handled.
- **Structure.** Sections come from Markdown/Word headings, with a fallback that infers headings from bold or large-font lines when no heading style is set. Spreadsheets and slide decks have no real heading semantics, so the thin-section check is naturally limited there.
- **Heuristics, not understanding.** The scanner flags *surface signals*. The genuinely hard judgements (re-deriving a figure, simulating a real challenge, confirming a hallucination) are delegated to the agent in Step 2 — by design.

Contributions of additional language packs and locale number formats are welcome.

---

## Project layout

```
anxiety/                          # Repo wrapper: docs + the shippable skill folder
├── README.md                     # This file
├── assets/
│   └── anxiety-mascot.svg        # README header illustration
├── .github/workflows/tests.yml   # CI: runs the test suite on Python 3.8–3.12
├── tests/
│   └── test_anxiety.py           # Stdlib unittest suite (extractor, parsers, every check, CLI)
└── anxiety/                      # The shippable skill (point your agent here)
    ├── SKILL.md                  # Skill definition and agent instructions
    ├── patterns/
    │   └── en.json               # English language pack (copy to add others)
    ├── references/
    │   ├── register-schema.md    # JSON register schema + category + checklist reference
    │   └── scanner-checks.md     # Full per-phase list of automated scanner checks
    └── scripts/
        ├── anxiety_scan.py       # Three-phase scanner (stdlib only)
        ├── extract_text.py       # Office/text extractor (stdlib only)
        └── validate_register.py  # Register schema/consistency validator (stdlib only)
```

## Requirements

- Python 3.8+ (standard library only — no `pip install` required).

## Tests

The suite is stdlib-only (`unittest`), mirroring the skill's no-dependency promise. It covers the
extractor, the number/currency parsers, every scanner check, the extraction-quality gate, the
diagnostic-only invariant, and CLI behaviour. CI runs it on Python 3.8–3.12.

```bash
python -m unittest discover -s tests -p "test_*.py"
```

## Trigger phrases

The skill is user-initiated. An agent activates it when you say things like:

> "run anxiety check" · "stress test this" · "what did I miss" · "poke holes in this" · "check my work" · "red-team this" · "is this defensible" · "would this survive scrutiny" · "prep for review"

## Author

Developed by **Arsalan Pardesi**.

## License

MIT © Arsalan Pardesi
