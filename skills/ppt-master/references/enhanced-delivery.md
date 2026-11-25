# Enhanced delivery support

Accepted local extension to PPT Master 6.9.0. Existing authorities and exporters
remain the implementation. No Anthropic source files are vendored.

## Choosing an expression from the page's job

Use this within the selected route's existing composition decision, not as an
extra phase, checklist or required planning artifact. Before carrying a deck's
visual motif onto a page, identify what its reader needs to notice, understand
or compare, and at what level of detail. Choose from the upstream expression
vocabulary according to that need and the usable information in the supplied
material. A coherent visual identity does not require a repeated representation.

When the useful expression is not obvious, consider meaningfully different ways
of making the important information perceptible, rather than only varying color
or decoration on the first layout. Select for both communication and visual
quality. The simplest whole view is valid when it does the job; variation and
object count are not goals. Do not invent missing evidence to enable a design.

At rendered review, check whether the chosen visual actually helps the reader
perform that page's job at the intended viewing size. If it only reinforces
identity while the relevant information remains inaccessible, revisit the
expression decision instead of merely polishing spacing or adding ornament.

## Final output is the visual review target

The upstream SVG checks stay mandatory. In addition, render the exported PPTX in
an available presentation renderer and inspect every delivered page. Prefer native
PowerPoint on Windows when installed, otherwise use an available LibreOffice
renderer. SVG/browser preview alone cannot verify the PPTX font substitution,
math rendering, native table layout or media compatibility. If no final-deck
renderer is available, report that limitation and do not claim visual verification.
Do not require a particular renderer to generate an otherwise valid deck.

Use `scripts/render_powerpoint.ps1` on Windows and the CLI documented in
`scripts/docs/enhanced-qa.md`. Open the source read-only; write separate outputs.
Never close or save another user presentation. Avoid spawning visible helper
windows. PowerPoint itself may display a window when COM requires one.

At presentation scale assess hierarchy, useful figure size, reading order,
whitespace, color consistency and deck rhythm. At detail scale inspect math,
labels, image crops, arrows, table values, clipping and unintended overlaps.
A purposeful diagram connection or background overlap is not automatically wrong.
Keep a compact per-page record of observations. Correct a demonstrated problem at
its owning source, export again and re-inspect affected pages. A shared style or
exporter change requires re-inspecting all affected pages. Stop when the requested
quality is achieved or disclose the concrete unresolved limitation; never manufacture
an automatic aesthetic pass from shape counts or a subjective score.

Use the existing `pptx_delivery_check.py` plus `enhanced_qa.py audit` for package
and native-object evidence. Counts prove presence, not fidelity or editability in
all apps. Label embedded paper figures as images; do not describe them as native
editable diagrams. Do not rasterize whole pages to make a preview look better.

## Precise edits

For an existing PPTX retain the upstream Edit Native route. Identify
the intended slide and object before modifying it. Save a new output file and
compare package parts plus object XML before/after using `enhanced_qa.py compare`.
For a single complete ordinary text run, this local extension permits
`scripts/native_text_patch.py` as a narrow Edit Native primitive instead of rebuilding
the target through roundtrip: require explicit slide, shape ID, exact old text and
new text. It refuses ambiguous/split runs and existing output paths. It preserves
the remaining XML bytes, including geometry, autofit, fonts and run formatting.
No text fitting is performed: inspect the result for overflow. All complex edits,
new objects, images, charts, notes and motion retain the original roundtrip route.
This explicit primitive exception does not create a fourth route.

On Windows, an import that encounters long generated preview paths should use a
short isolated workspace and copy the finished files back; never alter source
contracts or truncate source references to hide a path failure.

Declare allowed shape IDs, then check all remaining shapes, relationship files,
media, notes, masters and layouts. A scoped edit passes only if the requested change
is present and no undeclared change occurs. For changes intentionally affecting
several objects (e.g. a group move), list those objects rather than weakening the
comparison. Pure package serialization changes are reported separately; do not
misrepresent them as byte preservation. Render both versions, inspect the edited
area and verify the unchanged pages. A text replacement alone is not evidence
that layout edits, charts or animations all roundtrip correctly.

## Vector export

Keep editable SVG and PPTX sources. Export PDF through a presentation/vector
renderer and trim its page boxes without rasterizing. Select the intended figure
page and an explicit inspected bounding box in PDF points (`left bottom right top`).
Check the cropped PDF by rendering it and inspecting text/vector drawing operators
or embedded images. A PDF container may still contain raster figures. Disclose
which parts are vector, text outlines, native math or images. The crop tool never
infers a scientifically meaningful region. Do not claim a raster image inside PDF
became vector by export.

## Delivery and design latitude

Small visual ideas may improve spacing, color grouping, routing, emphasis and
auxiliary explanations, while preserving the algorithm, formulas, claims, values
and meaning of comparison images. Disclose the actual extra ideas briefly in the
conversation at delivery. No per-element scientific evidence ledger or strict
research-figure traceability workflow is required. Ordinary source/context notes
are sufficient unless the user asks otherwise.

For requested A/B evaluation use the same frozen content/assets and task prompt,
preserve both outputs and render them at equal resolution. Do not deliberately
weaken the baseline. State model/run differences, sample count and renderer limits.
Prefer concrete page observations to weighted scores; if scores help, label them
subjective. Never promise all future decks improve. Local installation and cloud
publication remain separate user decisions when requested by the user.
