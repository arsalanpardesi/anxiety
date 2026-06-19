# Scanner checks - what `anxiety_scan.py` catches automatically

Read this file when you need the full list of the scanner's pattern- and structure-based
checks per phase, which ones are advisory (off by default), and how the text is de-noised
before pattern matching. SKILL.md Step 1 summarizes this; the exhaustive list lives here.
For the machine-category names used in the register, see `register-schema.md`.

## Phase 1 - Coverage (what was missed?)
- Unreferenced source files - items in `--sources` whose subject never appears in the text.
- Thin sections - headed sections with too little substantive content.
- Missing expected sections - with `--checklist`, expected topics whose keywords never appear.

## Phase 2 - Correctness (is it actually right?)
- Mixed direction-of-change signals - sentences asserting both up and down for one metric.
- Extreme / implausible percentages - values >100% outside a growth/CAGR context (a percentage
  near growth cues like "grew", "rose", "CAGR", "year over year" is treated as normal, not an error).
- Inconsistent figures - the same metric stated with different values. Figures are keyed by the
  metric named nearest before them on the line (canonical `metric_aliases`, longest surface wins),
  so re-wordings of one metric are compared together while distinct metrics sharing a word
  ("revenue" vs "revenue growth") stay separate.
- Table / body mismatch - a labeled figure that differs between a table and the narrative.
- Post-cutoff data - years in the body that postdate the stated scope cut-off (needs `--context`).
- Mixed currency - two or more currencies used without a conversion note.
- Mixed measurement basis - e.g. CAGR vs YoY, nominal vs real, annual vs monthly in one document.
- Formula errors (`.xlsx`) - cells resolving to `#REF!` / `#DIV/0!` etc. (broken calculations).
- Hardcoded model cells (`.xlsx`) - typed-in numbers inside a formula row. *(advisory, `--advisory-checks`)*
- **Hallucination check** (see SKILL.md Step 2 - partly automated, must be completed by the agent).
- Numeric checks treat figures within 1% (and magnitude restatements like $5M vs $5,000,000) as equal, so rounding doesn't read as a conflict.

## Phase 3 - Adversarial (what would a hostile reader attack?)
- Deferred analysis - "should be analyzed", "requires further review", "was not available".
- Unqualified superlatives - "the highest", "the best" with no benchmark/comparator.
- Orphan assertions - quantified claims with no source/citation nearby. *(advisory, `--advisory-checks`)*
- Hedging-language density - "broadly", "approximately", "largely", etc. *(advisory, `--advisory-checks`)*
- Confidentiality / PII exposure - emails, phone numbers, credentials, keys. *(advisory, `--advisory-checks`)*

## Extraction quality
After a deliverable is extracted, a best-effort heuristic checks the text isn't garbled
(replacement characters, mostly non-letters, or lost word spacing — common when a PDF uses an
unsupported font/encoding). If it looks garbled the run sets `extraction_ok: false` in the
register, prints a warning, and the findings should be treated as unreliable until a
higher-fidelity reader is used (see SKILL.md). It never blocks the scan.

## Advisory checks and de-noising
Advisory checks are off by default because they misfire on real finance prose; enable them
with `--advisory-checks`. The remaining checks above run on every scan.

Pattern-based checks run on a de-noised copy of the text: fenced code blocks, inline code
spans, quoted spans ("..."), block quotes, and sections titled
"example/sample/illustrative/template/specimen" are stripped first so quoted or illustrative
text is not mistaken for the author's own claim. This is best-effort, not exhaustive.
