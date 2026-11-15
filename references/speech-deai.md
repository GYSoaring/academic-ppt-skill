# Speech De-AI Audit

This is an adaptation of the `lieflat-less-ai-tone` methodology for defense scripts. The original skill provides a full article-level detector and report pipeline; this adaptation keeps the parts that matter for spoken academic Chinese:

- script-detectable AI-tone patterns;
- spoken rhythm and breath groups;
- minimal edits that preserve facts, structure, and qualifications;
- an explicit "do not change" list.

It is not an authorship detector. It describes text features, not which model wrote the text.

## Two Layers

1. **Script layer** (`scripts/audit_speech.py`): counts patterns that can be expressed as regex or structure.
2. **Human/semantic layer**: metaphor quality, whether a three-item list is padding, and whether a transition is genuinely needed. The script reports candidates; the author decides.

The full original skill also has a semantic candidate pipeline and before/after reporting. For a defense script, the most valuable use is the pre-delivery audit: find the phrases that sound like a written report being read aloud, then rewrite only those.

## Levels

| Level | Meaning | Default action |
|---|---|---|
| A | Strong AI-tone rhetorical move | Rewrite the sentence |
| B | Empty metadiscourse or opener | Delete the frame |
| C | Translationese, dense list, or spoken-rhythm problem | Rewrite or split |
| F | Fingerprint-like punctuation or layout habit | Check per instance; not a verdict |
| X | Research shows no stable human/AI difference, or humans use it more | Do not change |

## A-Level: Rhetorical Reversal

Patterns:

- `不是 A，而是 B`
- `并非 A，而是 B`
- `与其说 A，不如说 B`
- `看似 A，实则 B`
- `你以为 A，其实 B`

Why it matters in a talk: the listener has to hold a rejected frame in working memory before getting the real claim. In a 45–90 second page, that is expensive.

Rewrite: state B first. Keep A only when the contrast is genuinely necessary, and put it in a separate sentence.

## B-Level: Empty Frame

Patterns:

- `一句话总结：`
- `核心是：`
- `关键在于：`
- `说白了，`
- `先说结论，`

Rewrite: delete the frame and say the content. If the first half carries a real transition, replace the colon with a comma or period instead of deleting it.

## C-Level: Translationese And Spoken Rhythm

### Long pre-modifiers

Break a noun phrase with more than 15 characters of modifiers into two clauses.

### `当……时`

Delete `当` and `时` when the time relation remains clear; let the clause stand as a front clause.

### `对于……来说` / `在……方面`

If the sentence is already about that object, delete the shell and make the object the subject.

### Sentence-initial signposts

`然而，` `因此，` `此外，` at the start of consecutive sentences sound like written paragraphs. Move the connector after the subject or delete it when the relation is obvious.

Do not delete them when the argument genuinely needs an explicit contrast.

### Dense enumeration

Three or more `、`-separated items in one breath are hard to follow. Group them, summarize them, or split the sentence. This is a speech-specific rule; a written list may legitimately keep all items.

### Semicolons

Replace `；` with `。` in spoken text. A semicolon asks the listener to hold a syntactic relation that they cannot see.

### Long sentences

The audit flags sentences over 70 Chinese characters. Split them into two breath groups. Do not mechanically alternate long and short sentences; that is "fake variance", not natural speech.

### Reading the slide

If a speech contains a slide bullet verbatim, the speaker is reading the slide. Rewrite the sentence as "this page answers X; the evidence is Y", not as a list of bullet phrases.

## X-Level: Do Not Change

The original paired study found no stable human/AI difference, or humans use the feature more, for:

- ordinal signposts such as `首先 / 其次 / 最后`;
- `这意味着` restatement;
- empty lead-in sentences followed by a list;
- summary substitution when the paragraph already contains the details;
- precise numbers;
- "pretty word" blacklists such as `赋能 / 闭环 / 抓手`;
- nominalization, simile, rhetorical questions, vague quantifiers;
- sentence-length variance as a goal.

Do not "de-AI" a defense script by deleting these. They are often normal academic speech.

## Metaphor And Anthropomorphism

The original skill treats metaphor quality as human review, not a model judgment. This adaptation does the same. If a metaphor is doing real explanatory work, keep it; if it is decorative, replace it with the mechanism it was trying to describe.

## Workflow

`build_speech.py` runs the audit automatically after writing `speaker_notes.md` and `defense_script.md`, producing:

- `speech_audit.md`: human-readable graded report;
- `speech_audit.json`: machine-readable counts and per-page hits.

Run it separately when editing an existing deck:

```bash
python scripts/audit_speech.py --deck deck.json --out work/speech_audit
```

Optional safe cleanup:

```bash
python scripts/audit_speech.py \
  --deck deck.json \
  --out work/speech_audit \
  --apply-safe \
  --out-deck deck.polished.json
```

`--apply-safe` only deletes known empty discourse markers and splits semicolons. It never rewrites an A-level argument, changes a fact, or invents a transition. Review the diff before delivery.

## Rewrite Discipline

- Minimal edit: one hit, one reason.
- Information conservation: names, numbers, dates, citations, conditions, and hedges stay.
- Structure preserved: slide order, headings, lists, and tables do not move.
- No forced "I/you" and no forced colloquialism; academic defense is allowed to sound formal.
- No authorship claim. "AI-tone" describes a pattern, not a person or a model.
- After rewriting, re-run the audit and compare the level counts; do not chase a zero by flattening the prose.

## Attribution And Licensing

Adapted from `lieflat-less-ai-tone` by the user's local skill. The original skill's full detector, semantic criteria, and before/after report are the authoritative implementation. If this adaptation is redistributed commercially, review the original skill's authorization note.
