# Academic deck review and validation

Read this reference after the first coherent deck pair exists and again before delivery.

## Evidence hierarchy

Keep these states distinct:

1. **Observed:** directly verified in the source paper, final files, full-size renders, or a local presentation application.
2. **Reported:** claimed by a builder or reviewer but not independently reproduced.
3. **Inferred:** a reasonable conclusion that still needs labeling.
4. **Pending:** an edit or check has started but has no final artifact and passing result yet.

Only observed evidence closes a validation gate. A reviewer saying “all checked” is a lead, not proof.

## Independent cross-review

Give the reviewer the paper, the locked brief, both final-candidate PPTX files, and the speaking limit. Ask for findings, not rewritten copy by default. Review these lanes separately:

- **Source accuracy:** numbers, cohorts, labels, methods, comparison groups, figure interpretation, and claim strength.
- **Academic reasoning:** research question, evaluation design, alternative explanations, limitations, and evidence boundary.
- **Defensibility:** likely questions, personal claims the presenter may not be able to support, and missing “why this paper” or “what next” reasoning.
- **Oral delivery:** natural phrasing, timing, transitions, long-number pronunciation, and an emergency skip path.
- **Visual quality:** template fidelity, hierarchy, density, blank-space balance, wrapping, overlap, crop quality, and legibility.
- **Artifact quality:** editability, notes separation, package integrity, fonts, relationships, and unfamiliar-machine risk.

Record each finding as `accept`, `rewrite`, or `reject`, with a short reason and affected slides. Rewriting is appropriate when the reviewer identifies a real gap but proposes artificial, inflated, or hard-to-defend language.

## Deterministic pair preflight

Run:

```bash
python3 scripts/check_deck_pair.py \
  --template "/absolute/path/template.pptx" \
  --submission "/absolute/path/submission.pptx" \
  --preparation "/absolute/path/preparation.pptx" \
  --expected-slides "${EXPECTED_SLIDE_COUNT}"
```

Set `EXPECTED_SLIDE_COUNT` from the locked brief. Use `--require-sources-slides` for preparation slides that contain externally sourced claims or assets. Use `--allow-empty-prep-slides` only for slides the brief intentionally leaves without spoken notes. Optional `--min-script-chars` and `--max-script-chars` are coarse bounds, not timing proof.

The script records final-file SHA-256 hashes and checks archive safety limits, ZIP CRCs, content-type declarations, relationship targets and linked/embedded content inventory, ordered and expected slide counts, dimensions, template-theme byte fidelity, visible text, object geometry, style, z-order, visibility, and recursively referenced media/chart payloads across the pair. It also checks submission-note relationship absence, preparation-note presence, orphan note parts, non-empty source blocks, and long numeric tokens in spoken notes. Preserve its hashes with the QA result so a later edit cannot inherit an earlier pass. Its JSON explicitly labels the result as a structural preflight and lists the manual gates still required; inherited layout/master fidelity and pixel-level equality remain part of the required render and template-fidelity review, along with semantic accuracy, font rendering, and real speaking time.

## Render and visual inspection

Render every slide of both variants after the final edit. Inspect each slide individually at full size; use a montage only for rhythm and consistency. Verify:

- both renders are visually identical slide by slide;
- no title wraps unexpectedly and no object clips, overflows, or overlaps;
- the copied template's master, layout, fonts, recurring bars, footers, and spacing remain intact;
- visual weight is balanced without low-value filler;
- figures, tables, axes, legends, and labels remain readable;
- all intended text and data objects are editable, while source images remain separately movable;
- no default or empty inherited placeholder survives in edit mode.

When exact image comparison is practical, compare corresponding rendered PNGs. A mismatch must be explained and authorized; notes alone should not alter the render.

## Content and speaking checks

- Recompute or re-read every important statistic from the source.
- Confirm that exact slide values and rounded spoken values refer to the same quantity and preserve the conclusion.
- Remove `[Sources]` blocks before estimating spoken length.
- Time at least one full read aloud at the presenter's realistic pace; leave buffer under the hard limit.
- Check that personal motivation uses only confirmed experience and survives a follow-up question.
- Confirm that limitations are stated once, naturally, and at the right strength; avoid repetitive defensive boilerplate.
- If the user requests a closing slide, make it consistent with the template and keep it free of unrequested slogans or filler.

## Package and local-application checks

- Confirm the PPTX is a valid ZIP and every packaged part has a content-type declaration.
- Confirm every internal OOXML relationship resolves. Stray theme or media parts can trigger repair warnings on stricter Office-compatible applications.
- Keep template-theme byte preservation separate from effective visible-theme comparison. A notes-master-only theme change violates a strict template package contract but does not by itself prove that slide renders changed.
- Open the submission deck in a local PowerPoint-compatible application, move quickly through all slides, enter slide-show mode, select representative text, and select or edit a representative chart/table/shape.
- Open the preparation deck, confirm notes appear on every required slide, and confirm source blocks are visually separated from the script.
- Confirm the submission deck contains no speaker script or hidden preparation content.
- If the deck will move to an unfamiliar machine, test the copied file there when possible and consider an explicitly requested PDF fallback.

Any new page, reordered slide, note rewrite, or visible copy change invalidates the prior pair result. Rebuild both variants and rerun the complete gate.
