---
name: academic-presentation-deck
description: Turn academic papers and user-supplied presentation templates into editable, source-checked submission and preparation decks with speaker notes, speaking constraints, independent cross-review, and package and visual validation. Use for academic interviews, paper presentations, defenses, seminars, and conference talks (学术面试、论文汇报、答辩); do not use for generic business or sales decks.
---

# Academic Presentation Deck

Create a deck the user can submit, edit, present naturally, and defend under questioning. This skill adds an academic artifact contract to the available presentation-authoring workflow; follow the presentation skill's template, authoring, rendering, and delivery requirements whenever a PPTX is created or edited.

## Lock the brief

Before authoring, resolve from the request, local files, and conversation history:

- the paper and any supplementary material that serve as factual authority;
- the exact template or visual reference, if supplied;
- audience, occasion, purpose, language, hard time limit, and expected question depth;
- required slide count, filenames, submission rules, and whether both deck variants are wanted;
- typography, bilingual treatment, signature/footer rules, and explicit visual dislikes;
- the user's real preparation level and experiences that may safely appear in the talk.

Write one internal communication job: by the end, the audience should understand the paper's defensible contribution, evidence, limitations, and why the presenter chose it. Ask only when a missing choice would materially change the result and cannot be recovered locally. Mark ambiguous or reconstructed source material plainly; never invent a paper fact, credential, contribution, personal experience, or publication status.

## Preserve the template and editability

Treat a user-designated PPTX as a contract, not inspiration. Inspect every source slide and use the presentation skill's template-following route: preserve the master-layout-slide hierarchy, reuse inherited layouts and elements, and retain the template's fonts, spacing, title treatment, and recurring chrome unless the user asks otherwise.

Prefer native editable text, tables, charts, lines, and simple shapes. Keep source figures as movable images when rebuilding them would reduce fidelity. Do not flatten a whole slide into an image. Do not inject generic deck mannerisms, decorative filler, slogans, extra disclaimers, or a new font system. Concision must not become an empty slide: use evidence-bearing figures, tables, or structure when they advance the academic argument.

## Build the academic narrative

Choose a cumulative arc suited to the event. A common interview or defense arc is research question, data and design, method, validation, results, evidence boundary, limitations, the presenter's interpretation, and likely next study. Do not turn a research deck into popularization or a catalog of paper sections.

- Give each slide one academic job and one primary claim.
- Cross-check every displayed number, sample definition, comparison, and limitation against the paper.
- Distinguish observed results, scenario calculations, post-hoc analyses, and causal claims.
- Preserve exact values on the slide. In speaker notes, use accurate rounded phrasing when exact long numbers would impede delivery.
- Surface the evaluation choices and evidence boundaries most likely to attract questions.
- Write natural audience-facing copy. Remove repetitive qualifiers, mechanical parallelism, and sentences a presenter would not plausibly say.

## Produce one visible deck and two delivery variants

Unless the user requests a different contract, generate both variants from the same visible content source:

1. **Submission deck:** editable visible slides with no speaker-note script or hidden preparation text.
2. **Preparation deck:** the same editable visible slides, with a per-slide spoken script and a separate `[Sources]` block in notes where claims or assets require provenance.

The two variants must have the same slide order, visible wording, layout, and rendered appearance. A change to any visible slide must propagate to both variants before validation. Do not count `[Sources]` text as spoken material, and do not instruct the presenter to read it aloud.

## Write notes for speaking, not reading

Allocate the time budget across slides, leaving room for transitions and interruption. Draft notes in the user's natural register and at a depth they can defend. Personal motivation must be grounded in confirmed experience; reject or rewrite reviewer language that inflates deployment experience, research ownership, or expertise.

Use exact timed rehearsal as the completion test. Character or word counts are estimates only. Keep the visible exact number and the spoken rounded number semantically consistent. For a hard-stop setting, add a private skip path or transition in the preparation notes rather than exposing timing scaffolds on the slide.

## Cross-review without surrendering judgment

After a coherent first build, run an independent review against the paper, brief, and actual deck files. Require evidence for findings and separate factual errors, overclaims, oral-delivery risks, visual defects, and package defects. Classify each suggestion as accept, rewrite, or reject; do not paste a reviewer's prose merely because its direction is useful. Revalidate both variants after every accepted change.

Before review or delivery, read [references/review-and-validation.md](references/review-and-validation.md). Use `scripts/check_deck_pair.py` as the deterministic pair/package preflight; it supplements rather than replaces full-slide rendering and human visual inspection.

## Completion gate

Do not call the task complete while a rebuild is still running or because an authoring command exited successfully. Completion requires final files at their promised paths, successful package and pair checks, every slide rendered and inspected at full size, no unresolved overflow or placeholder defects, confirmed note separation, and a local PowerPoint-compatible application spot-check when available. For high-stakes transfer to an unfamiliar machine, mention a PDF fallback as an option; create it only when requested or included in scope.
