---
name: anxiety
description: "Systematic re-interrogation of any completed knowledge-work deliverable (report, memo, analysis, brief, proposal, research write-up) across three phases: Coverage (what was missed?), Correctness (is it actually right?), and Adversarial (what would a hostile reader attack?). Catches hallucinated access claims, unsourced figures, deferred analysis, and thin coverage. Produces a structured Anxiety Register in JSON and markdown. Diagnostic only, never edits the deliverable. User-initiated."
license: MIT
metadata:
  version: 0.1.0
  author: Arsalan Pardesi
  compatibility: "Requires Python 3.8+ (standard library only) and a shell to run the bundled scanner; falls back to manual agent checks when no Python is available."
---

# Skill: Anxiety Check — Re-Interrogate Completed Work

Use this skill when the user asks to:

- "run anxiety check", "stress test this", "what did I miss", "poke holes in this"
- "check my work", "am I missing something", "review for completeness"
- "challenge the analysis", "red-team this", "adversarial review", "pre-mortem"
- "is this defensible", "would this survive scrutiny", "prep for review"

It works on **any completed knowledge-work deliverable**: a consulting report, research
write-up, legal memo, strategy doc, due-diligence report, technical design, grant proposal,
investment memo, board paper, or academic paper.

**Supported file types** (the usual places people work):

| Type | Extensions | How it is read |
|---|---|---|
| Word | `.docx` | Headings, paragraphs, and tables extracted to text |
| Excel | `.xlsx` | One section per sheet; rows read as table lines |
| PowerPoint | `.pptx` | One section per slide; slide title becomes a heading |
| PDF | `.pdf` | Text layer extracted from content streams (best-effort) |
| Markdown / text | `.md` `.markdown` `.txt` `.rst` | Read as-is |

Extraction is **stdlib-only** (Office files are ZIP+XML parsed with `zipfile` + `xml.etree`;
PDF content streams are decoded with `zlib`) so there are still no third-party dependencies.
It's best-effort by design: the Office reader captures headings, tables, and slide text but
can miss complex layouts, charts, embedded objects, comments, or tracked changes; the PDF
reader undoes the Flate/LZW/ASCII85/ASCIIHex/RunLength filters but **cannot** read
scanned/image-only PDFs (no OCR), encrypted PDFs, predictor-encoded or image streams, or
custom/CID font encodings; and legacy binary formats (`.doc`, `.xls`, `.ppt`) aren't parsed
directly. Whenever the harness offers a higher-fidelity reader for the file at hand, prefer
it (see the workflow note above); otherwise the bundled extractor is the always-available
fallback, or save/export to text/markdown first.

This skill is **user-initiated only**. It never runs automatically and it is purely
**diagnostic** — it identifies and reports issues; it does not edit the deliverable unless
the user explicitly asks for a fix afterward.

## Core idea

Most review tooling checks whether a document is *well-formed* (headings, formatting, length).
The anxiety check interrogates whether the **underlying work is thorough, correct, and
defensible**:

| Question | Phase |
|---|---|
| Were all available sources actually used? | Coverage |
| Do the stated figures and claims hold up? Did we invent anything? | Correctness |
| Would a hostile, expert reader be able to attack it? | Adversarial |

A formatting review asks "is this written well?" The anxiety check asks "is this **true** and
**complete**, and will it **survive challenge**?"

## Inputs

Required:
- The deliverable to interrogate — a Word, Excel, PowerPoint, PDF, markdown, or text file
  (or pasted text).

Optional (improves the Coverage phase):
- A **sources directory** — folder of files that *should* have informed the deliverable
  (notes, source documents, data exports, transcripts). The scanner flags source files whose
  topic never appears in the deliverable.
- A **scope/context note** — what access actually existed (e.g. "desk research only, no
  interviews", "based on Q3 data only"). This sharpens the hallucination check.

## Execution

### Step 1 — Run the portable scanner (optional pattern-based pre-pass)

The bundled scanner is **stdlib-only Python** (no third-party dependencies, no database, no
network) so it runs in any harness with a shell. Treat it as an **accelerant, not the
detector**: it catches structural and pattern issues cheaply and reproducibly, but the
findings that matter most on real analytical work come from the agent reasoning in Step 2.
A fast review can skip straight to Step 2; the scanner just front-loads the mechanical checks.

Some pattern checks (orphan assertions, hedging density, PII/secret patterns) have a high
false-positive rate on real finance prose, so they are **off by default**. Add
`--advisory-checks` to include them when you specifically want that sweep.

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

Optional flags:
- `--checklist` runs a **coverage checklist** (a path to a JSON file) and flags expected
  sections whose keywords never appear. The checklist is generated per deliverable in Step 1b
  below rather than shipped, so it fits whatever this particular piece of work claims to be.
- `--baseline` diffs this run against a prior register and writes `<name>.anxiety.diff.md`
  (which checks were resolved, which are new), so you can re-run after fixes and see progress.

The language-specific cues (hedges, deferred-analysis phrases, hallucination patterns, etc.)
live in pattern packs under `patterns/`. English ships built-in and as `patterns/en.json`.
To analyse a deliverable in another language, copy `patterns/en.json`, translate the cue
lists, save it as `patterns/<code>.json`, and pass `--lang <code>` (or point `--patterns` at
any pack file). If a pack is missing or unreadable the scanner falls back to English and
prints a note, so it never hard-fails.

The scanner auto-detects the file type and extracts the text before analysing it. Word,
Excel, PowerPoint, and PDF files are handled natively (PDF text-layer extraction undoes the
Flate/LZW/ASCII85/ASCIIHex/RunLength filters). To inspect just the extracted text (useful
for debugging) run `python scripts/extract_text.py <file> [out.md]`.

**Prefer the harness's own parsers when they're better — for any file type, not just PDF.**
The bundled extractor is stdlib-only and deterministic, which makes it portable, but it isn't
the best reader for every file. For Word/Excel/PowerPoint it captures headings, tables, and
slide text but can miss complex layouts, charts, embedded objects, comments, or tracked
changes; for PDF it can't read scanned/image-only pages (no OCR), encrypted files,
predictor-encoded or image streams, or custom/CID font encodings. So before falling back to
the bundled extractor, check whether the host environment offers a higher-fidelity reader and
use it first:

- **Word / Excel / PowerPoint** — native file reading, `pandoc`, `markitdown`, LibreOffice
  (`soffice --convert-to`), `python-docx`/`openpyxl`/`python-pptx`.
- **PDF** — native file reading, `pdftotext`/poppler, `pandoc`, `markitdown`,
  `pdfminer`/`pypdf`, or an OCR tool for scanned pages.
- **Anything else** — whatever converter the harness provides.

When you use one of these, **save its output to a `.md`/`.txt` file and run the scanner on
that file**, and record which extractor was used in the register (e.g. in the `context`/notes)
so the findings stay traceable. Fall back to the bundled extractor whenever no better tool is
available; never assume one exists. The extracted text is the input that drives every check,
so a cleaner extraction directly improves the review.

Outputs (written next to the deliverable, or to `--out`):
- `<name>.anxiety.json` — structured register with every finding
- `<name>.anxiety.md` — human-readable summary

If no Python is available, the agent performs the equivalent checks manually using the same
methodology (see the category list below) — the scanner is an accelerant, not a hard
dependency.

**What the scanner catches automatically:** structural coverage gaps (unreferenced sources,
thin sections, missing checklist sections), correctness signals (inconsistent figures, table/body
mismatches, post-cutoff data, mixed currency or measurement basis, `.xlsx` formula errors), and
adversarial cues (deferred analysis, unqualified superlatives). Some checks (orphan assertions,
hedging density, PII/secret patterns, hardcoded model cells) are advisory and off by default; add
`--advisory-checks` to include them. Pattern checks run on a de-noised copy of the text (fenced
code, inline code, quoted spans, block quotes, and example/sample/illustrative sections are
stripped first) so quoted or illustrative text is not mistaken for the author's own claim.

The full per-phase catalogue is in `references/scanner-checks.md`; the machine-category list is in
`references/register-schema.md`.

### Step 1b — Build a coverage checklist for this deliverable (optional)

Coverage gains a lot from knowing what this specific deliverable was *supposed* to contain. Do
not rely on a fixed library of domain checklists — none ships, because a small set implies a
narrow scope and goes stale. Instead **generate a checklist tailored to this deliverable**, then
feed it to `--checklist`.

1. Work out the deliverable's **type, audience, and the decision it supports** from the document
   itself plus the `--context` note (e.g. "quality-of-earnings report for an IC, supports a buy
   decision"; "market study for a product team"; "legal memo for the GC").
2. If that is genuinely unclear, ask the user 3–6 short intake questions: what kind of
   deliverable is this; who is the audience and what decision does it support; what must it
   contain to be considered complete; what is explicitly out of scope; any house template or
   standard it must follow.
3. Emit a checklist object of the shape the scanner consumes:

   ```json
   {
     "name": "Quality of Earnings",
     "sections": [
       {"title": "EBITDA normalizations", "severity": "high",
        "keywords": ["normaliz", "add-back", "one-off", "non-recurring"]},
       {"title": "Net working capital", "severity": "high",
        "keywords": ["working capital", "nwc", "peg"]}
     ]
   }
   ```

   Keep `keywords` lowercase substrings (matched case-insensitively); a section with no keywords
   can never match and is skipped. **Set severity by impact, not habit:** `high` only if the
   section's absence would change the conclusion or block the decision; `medium` for
   expected-but-not-load-bearing; `low` for nice-to-have.
4. Write it next to the deliverable as `<name>.anxiety.checklist.json`.
5. Run the scanner with `--checklist <name>.anxiety.checklist.json`. Missing sections surface as
   `missing_expected_section` coverage findings.

This keeps the tool credible across *any* knowledge work: the bar is whatever this deliverable
set for itself, not a generic archetype.

### Step 2 — Agent-mediated deep checks (MANDATORY)

The scanner catches structure and patterns. The agent MUST add the checks that require
reasoning. Append each finding to the register (Step 3).

**Two rules apply to every agent finding in this step:**

1. **Ground it or label it.** Each finding MUST quote the specific text it is about (a verbatim
   line, cell, or sentence from the deliverable) in `evidence`, plus the source it was checked
   against. If no source was available to verify the claim, you MUST NOT assert a discrepancy —
   instead record it as severity `low` with the title prefixed `Unverifiable —` and state in
   `detail` that no source was provided. This prevents the review from inventing its own findings.
   **For any numeric finding, do not rely on mental arithmetic** — see the compute-tool rule in 2b.
2. **Prioritize, don't cap.** Cover **all material** figures, claims, and conclusions, ordered by
   impact on the conclusions. The numbers below are minimums, not ceilings: on a short deliverable
   they may be the whole set; on a long one, keep going until the material items are exhausted.

**2a. Figure / claim re-verification (the most material; at least 5)**
For each load-bearing figure or factual claim, working from most to least material:
1. Locate the source that should support it (a file in `--sources`, a dataset, a citation).
2. Re-derive or re-read the value independently.
3. Compare to what the deliverable states. Flag any discrepancy >1% (or any factual mismatch)
   with the correct value, quoting both the deliverable text and the source.

**2b. Internal arithmetic recomputation (do this even with no `--sources`)**
This is the highest-value check when no source workbook is provided, because it needs none:
recompute the document's own math from the numbers it already states, and flag anything that
doesn't tie out.

**Use a real compute tool — never mental math.** Where the harness offers code execution, a
Python/shell interpreter, a calculator (`bc`), or the spreadsheet/model itself, use it to do
every calculation, and quote the exact expression and inputs you ran in `evidence`
(e.g. `EV = equity 18.2 + net debt 4.4 = 22.6 ≠ stated 24.1`). Reason manually only when no
compute tool exists, and say so in the finding. Mental arithmetic on a precision check defeats
its purpose: it both misses real errors and invents false ones.

Quote the figures you used in `evidence`. Check at least:
- **Totals vs parts** — do segment / line-item figures sum to the stated total (revenue by
  segment, sources & uses, bridge waterfalls)?
- **Build-ups** — does the stated total reconcile to its components (e.g. EV = equity value +
  net debt; net debt = gross debt − cash; equity value = share price × shares)?
- **Ratios & margins** — does each stated margin/ratio equal its inputs (gross margin =
  gross profit ÷ revenue; net debt / EBITDA)?
- **Growth rates** — does a stated CAGR match the endpoints `(end/start)^(1/years) − 1`, and
  do period-over-period % changes match the underlying values?
- **Restatements** — is the same quantity (e.g. net debt) defined and stated consistently
  everywhere it appears?
Flag any mismatch beyond rounding with the recomputed value and the inputs you used.

**2c. Challenge simulation (the most important conclusions; at least 3)**
For each important conclusion, working from most to least important:
1. Write the strongest counter-argument an expert opponent would make.
2. Check whether the deliverable pre-empts it.
3. If not, draft a sentence that addresses the challenge.

**2d. Cross-document reconciliation (if multiple related deliverables exist)**
Extract the key shared figures/claims across the documents and verify each is stated
consistently. Flag discrepancies and name which document should be authoritative.

**2e. Anchoring / single-source risk**
For figures or claims central to the conclusions, check whether they rest on a single
unverified source. Flag any material claim that was never cross-validated.

**2f. Hallucination check (CRITICAL — always perform)**
Detect fabricated context that makes AI-generated work sound more authoritative than the
evidence supports. Cross-reference the deliverable against the actual scope/context (the
`--context` note, scope section, or what the user states).

- **False access claims** — "interviews were conducted", "site visit", "we spoke with",
  "management confirmed", "expert call on <date>", "survey of N respondents". If no such
  access occurred, every instance is a hallucination.
- **Fabricated citations** — references to documents, pages, sections, or studies that do not
  exist in the provided sources. Spot-check ~10 citations against the source set.
- **Invented data / false precision** — suspiciously specific statistics ("increased 87% in
  2021–2022") attributed to named authorities (Gartner, McKinsey, a journal) that cannot be
  traced. Generic estimated ranges marked as estimates are acceptable; false-precision
  numbers attributed to a named source must be verifiable.
- **Phantom capabilities** — "our testing confirmed", "physical inspection revealed", "we
  audited the system" when no such activity took place.
- **Temporal / engagement inconsistencies** — data or events postdating the stated cut-off;
  claimed durations, team members, or interactions that did not happen.

Severity for hallucination findings:
- **Critical** — false claims of access/interaction/testing that misrepresent the evidence
  basis (e.g. "10 interviews were conducted" when none occurred). If the document's own scope
  statement says "no interviews / desk research only" but the body claims interviews, this is
  an **automatic Critical**.
- **High** — citations pointing to non-existent sources.
- **Medium** — invented statistics with false precision attributed to a named source.
- **Low** — phrasing that mildly implies capability beyond what occurred.

### Step 3 — Merge agent findings into the register

Append each agent finding to the JSON register using the standard entry shape (every finding
must carry grounding `evidence` per the rule above), then regenerate the markdown summary
(re-run the scanner's writer, or edit both files consistently).

**Record that you ran the Step-2 checks.** The register seeds an `agent_checks` block with one
entry per reasoning check (`figure_re_verification`, `internal_recompute`, `challenge_simulation`,
`cross_document`, `single_source_risk`, `hallucination`), each starting as `"pending"`. As you
complete each check, set its `status` to `"completed"` (or `"not_applicable"` with a short
`note` saying why), and set `items_checked`. This is the honesty ledger: the scanner cannot do
these checks, so a register full of green automated phases means nothing until these are filled in.

After editing, validate the register so a malformed entry cannot silently corrupt the artifact:

```bash
python scripts/validate_register.py <name>.anxiety.json
# before presenting, gate on the reasoning checks actually being done:
python scripts/validate_register.py <name>.anxiety.json --require-complete
```

The validator checks required keys, enum values, id uniqueness, and that the summary scores
match the findings. With `--require-complete` it also fails if any `agent_checks` entry is still
`pending`. Fix any reported error before presenting results.

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

### Step 4 — Present results to the user

Summarize concisely, leading with the verdict:
- The **verdict band** (Not ready / Needs work / Solid — minor items / Clean) and the count of
  blocking (Critical/High) vs advisory (Medium/Low) findings.
- The severity-weighted score (0–100) and the three phase scores, as a secondary signal.
- The top 3 most actionable findings with specific remediation steps.

## Register entry shape

For the full top-level register structure and the complete machine-category list (including
the agent-only categories), read `references/register-schema.md`.

```json
{
  "id": "ADV-003",
  "phase": "coverage | correctness | adversarial",
  "category": "<machine category, e.g. orphan_assertions>",
  "severity": "critical | high | medium | low",
  "title": "<short headline, lead with the count where relevant>",
  "detail": "<what the issue is and why it matters>",
  "evidence": "<line refs, quotes, or source pointers>",
  "remediation": "<specific, actionable fix>",
  "auto_fixable": false,
  "status": "open"
}
```

## Scoring

**The verdict is the headline, not the number.** A single additive percentage is misleading
on a thorough review: enough minor findings will sink any substantial document, which punishes
completeness. So the register leads with a **verdict band** driven by the worst open severity:

| Verdict | Trigger |
|---|---|
| **Not ready** | any open **Critical** |
| **Needs work** | any open **High** (no Critical) |
| **Solid — minor items** | only Medium/Low open |
| **Clean** | no findings |

A severity-weighted **score (0–100)** is still reported for continuity, but it is built to not
collapse from volume: **blocking** findings (Critical/High) deduct in full (−25 / −15), while
**advisory** findings (Medium/Low) deduct −8 / −3 only up to a cap (24 per phase, 36 overall).
A long tail of low-severity, heuristic items therefore can't drown out a sound document.
Dismissed findings don't count toward either the score or the verdict, so triaging false
positives is rewarded. Read the score as a rough severity-weighted signal; act on the verdict
and the blocking count.

## Severity anchors

Use these anchors so severity is consistent run-to-run (the hallucination rubric in Step 2e
overrides for that category):

| Severity | Use when… | Typical categories |
|---|---|---|
| **Critical** | the work misrepresents its evidence basis, or a core conclusion is wrong | hallucinated access/interaction (scope says none), a load-bearing figure that doesn't reconcile |
| **High** | a material figure/claim is unsupported or inconsistent, or a key challenge is unaddressed | `inconsistent_figures`, `table_body_mismatch`, fabricated citation, failed internal recomputation on a headline number, `thin_sections` (>3) |
| **Medium** | a real but non-load-bearing issue, or a likely-but-unconfirmed problem | `mixed_direction_signals`, `extreme_percentages`, `post_cutoff_data`, `mixed_currency`, `deferred_analysis` |
| **Low** | minor, stylistic, or heuristic-only; or unverifiable for lack of a source | `mixed_baseline`, `unqualified_superlatives`, any finding titled `Unverifiable —` |

## Diagnostic only — no automatic fixes (ABSOLUTE RULE)

The anxiety check is a **diagnostic tool**, not a repair tool.

- The scanner MUST NOT modify the deliverable, the sources, or any user work product. (It only
  writes its own `*.anxiety.json` / `*.anxiety.md` register files.)
- The agent MUST NOT apply fixes during Steps 1–4 unless the **user explicitly asks** to fix a
  specific item or category.
- After presenting results, wait for instruction. If the user says "fix the orphan
  assertions" or "resolve finding #3", perform that specific fix. Otherwise, change nothing.
- Never chain anxiety-check → auto-fix.

**Correct flow:** run check → present findings → user requests a fix → apply that fix → optionally re-run to show improvement.

**Prohibited flow:** run check → silently rewrite the deliverable.

## What this skill does NOT do

- It does not second-guess legitimate professional judgment calls.
- It does not rewrite the deliverable — it flags issues for the user to decide on.
- It does not replace a domain expert's substantive review — it complements it with
  systematic coverage, correctness, and adversarial pressure.
- It is not part of any automated pipeline — it is user-initiated only.
