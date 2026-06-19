# Anxiety Register - schema reference

The scanner and the agent both write to the same JSON register so findings from the
automated pass and the agent-mediated deep checks live together.

Read this file when you need the exact register structure, the finding shape, or the full
list of machine categories (scanner-produced and agent-produced).

## Contents

- [Top-level object](#top-level-object)
- [Finding object](#finding-object)
- [Categories](#categories)
- [Checklist file (`--checklist`)](#checklist-file---checklist)
- [Agent checks (completeness ledger)](#agent-checks-completeness-ledger)
- [Scoring](#scoring)

## Top-level object

```json
{
  "anxiety_register_id": "uuid",
  "deliverable_path": "path/to/deliverable.docx",
  "deliverable_format": "Word (.docx) | Excel (.xlsx) | PowerPoint (.pptx) | PDF (.pdf) | Markdown | Plain text",
  "language": "en",
  "context": "stated scope note, e.g. 'desk research only; no interviews'",
  "timestamp": "ISO-8601 UTC",
  "tool": {
    "name": "anxiety",
    "version": "0.1.0",
    "advisory_checks": false,
    "checklist": "Quality of Earnings | null"
  },
  "phases": {
    "coverage":    { "score": 0-100, "findings": [ <finding>, ... ] },
    "correctness": { "score": 0-100, "findings": [ <finding>, ... ] },
    "adversarial": { "score": 0-100, "findings": [ <finding>, ... ] }
  },
  "agent_checks": {
    "figure_re_verification": { "status": "pending | completed | not_applicable", "items_checked": 0, "note": "" },
    "internal_recompute":     { "status": "pending", "items_checked": 0, "note": "" },
    "challenge_simulation":   { "status": "pending", "items_checked": 0, "note": "" },
    "cross_document":         { "status": "pending", "items_checked": 0, "note": "" },
    "single_source_risk":     { "status": "pending", "items_checked": 0, "note": "" },
    "hallucination":          { "status": "pending", "items_checked": 0, "note": "" }
  },
  "summary": {
    "verdict": "not_ready | needs_work | review_minor | clean",
    "verdict_label": "human-readable verdict",
    "total_findings": 0,
    "blocking": 0, "advisory": 0,
    "critical": 0, "high": 0, "medium": 0, "low": 0,
    "auto_fixable": 0,
    "coverage_score": 0, "correctness_score": 0,
    "adversarial_score": 0, "overall_confidence": 0
  }
}
```

## Finding object

```json
{
  "id": "COV-001 | COR-002 | ADV-003 | <PHASE>-AGENT-001",
  "phase": "coverage | correctness | adversarial",
  "category": "machine category, e.g. orphan_assertions",
  "severity": "critical | high | medium | low",
  "title": "short headline; lead with the count where relevant",
  "detail": "what the issue is and why it matters",
  "evidence": "line refs, quotes, or source pointers",
  "remediation": "specific, actionable fix",
  "auto_fixable": false,
  "status": "open | resolved | dismissed"
}
```

Use `<PHASE>-AGENT-NNN` ids for findings added by the agent in Step 2 so they are
distinguishable from scanner findings.

## Categories

| Phase | Category | Source |
|---|---|---|
| coverage | unreferenced_sources | scanner (needs `--sources`) |
| coverage | thin_sections | scanner |
| coverage | missing_expected_section | scanner (needs `--checklist`) |
| coverage | unanswered_questions | agent |
| correctness | mixed_direction_signals | scanner |
| correctness | inconsistent_figures | scanner |
| correctness | extreme_percentages | scanner |
| correctness | table_body_mismatch | scanner |
| correctness | post_cutoff_data | scanner (needs `--context` with a year) |
| correctness | mixed_currency | scanner |
| correctness | mixed_baseline | scanner |
| correctness | formula_errors | scanner (`.xlsx`) |
| correctness | hardcoded_in_formula_range | scanner (`.xlsx`, advisory, `--advisory-checks`) |
| correctness | hallucination | scanner + agent |
| correctness | figure_re_verification | agent |
| correctness | internal_recompute | agent |
| correctness | cross_document | agent |
| correctness | single_source_risk | agent |
| adversarial | deferred_analysis | scanner |
| adversarial | unqualified_superlatives | scanner |
| adversarial | orphan_assertions | scanner (advisory, `--advisory-checks`) |
| adversarial | hedging_language | scanner (advisory, `--advisory-checks`) |
| adversarial | confidentiality_pii | scanner (advisory, `--advisory-checks`) |
| adversarial | challenge_simulation | agent |

Pattern-based checks run on a de-noised copy of the text: fenced code blocks, inline code
spans, quoted spans ("..."), block quotes, and sections whose heading matches
example/sample/illustrative/template/specimen are removed first so quoted or illustrative
text is not mistaken for real claims. Structural checks (thin sections, unreferenced sources)
run on the original text.

Numeric checks (`inconsistent_figures`, `table_body_mismatch`) treat values within 1% as equal,
and normalize magnitudes (so `$5M` and `$5,000,000` match), so rounding is not flagged as a
conflict.

## Checklist file (`--checklist`)

Coverage checklists are **generated per deliverable** (usually by the agent in SKILL.md
Step 1b) and passed to `--checklist <path>`; none ship with the skill, so the bar fits the
specific deliverable rather than a generic archetype. A checklist is a JSON file of this shape:

```json
{
  "name": "human-readable label, e.g. 'Quality of Earnings'",
  "sections": [
    {
      "title": "expected section name",
      "severity": "critical | high | medium | low",
      "keywords": ["lowercase", "substrings", "matched case-insensitively"]
    }
  ]
}
```

Each section whose `keywords` never appear in the deliverable produces a
`missing_expected_section` coverage finding at the section's `severity`. Set severity by impact:
`high` only when the section's absence would change the conclusion or block the decision,
`medium` for expected-but-not-load-bearing, `low` for nice-to-have. A section with an empty
`keywords` list can never match and is skipped. Loading is fail-soft: a missing or malformed
file warns and the run continues without checklist coverage. By convention the generated file is
saved as `<name>.anxiety.checklist.json` next to the deliverable.

## Agent checks (completeness ledger)

`agent_checks` records whether the Step-2 reasoning checks were actually performed. The scanner
seeds every entry as `"pending"`; the agent sets each to `"completed"` or `"not_applicable"`
(with a `note`). `validate_register.py --require-complete` fails while any entry is still
`pending`, so a register cannot be presented as done before the reasoning checks are done.

The `tool` block records the tool version and the run's options (`advisory_checks`, `checklist`)
so a register is self-describing. Re-running with `--baseline <prior>.anxiety.json` writes a
`<name>.anxiety.diff.md` summarizing which finding categories were resolved or newly introduced.

## Scoring

The headline is the `verdict`, driven by the worst open severity:

| Verdict | Trigger |
|---|---|
| not_ready | any open Critical |
| needs_work | any open High (no Critical) |
| review_minor | only Medium/Low open |
| clean | no findings |

The severity-weighted `*_score` / `overall_confidence` values (0-100) deduct blocking findings
(Critical -25, High -15) in full and advisory findings (Medium -8, Low -3) up to a cap
(24 per phase, 36 overall), floor 0. Dismissed findings are excluded from both the scores and
the verdict.
