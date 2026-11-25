# Finished-file checks in the local enhancement

These original helpers supplement the existing SVG and delivery checkers. They
do not replace the existing authoring, native editing, formula, chart, or export
routes. Run them with the Python runtime that has PPT Master's dependencies;
PDF cropping additionally needs `pypdf`.

## Finished PPTX audit

```powershell
python scripts/enhanced_qa.py audit deck.pptx --report validation/finished.json
```

Checks ZIP CRC, duplicate and noncanonical parts, XML parsing, internal
relationship targets, required parts, slide roster, duplicate shape IDs, native
object counts, and geometric bounds concerns. `--strict-bounds` returns nonzero
for those concerns as well as integrity errors. Without it, `review-needed`
means the file needs inspection, not that it is broken. Inspect intentional
off-canvas elements before changing them. Group rotation, text wrapping, ink
bounds, effects, and inherited placeholder geometry need native renders.

Counts prove the presence of native objects, not that every visible mark is
editable. A picture of a formula is still a picture. Read the formula count and
inspect the delivered deck in PowerPoint when formula editability matters.

## Exact local edit comparison

```powershell
python scripts/enhanced_qa.py compare before.pptx after.pptx --allow-shape 4:27 --report validation/edit.json
```

The slide number is the 1-based presentation order; the ID is `p:cNvPr/@id`.
Repeat `--allow-shape` for each explicitly requested native object. Use the
smallest needed object: allowing a group permits the whole group subtree.

Every unapproved ZIP member payload must be byte-for-byte identical. Within an
approved slide, removing only the approved native shape subtrees must leave
canonical XML identical, including order and non-shape properties. Extra shape,
relationship, media, notes, or metadata changes fail with a nonzero exit code.
ZIP compression bytes and headers can differ when rebuilding the archive; the
check concerns all uncompressed member payloads. Shape XML that remains identical
can still be rendered differently by another program: inspect both renders.

This strict checker is for a bounded local edit, not an automatic authorization
for package-wide edits. If the requested edit needs a new image relationship or
another part, identify that scope before editing and use the existing route's
corresponding checks; do not broaden the shape allowlist to hide the change.

## Native PowerPoint rendering (Windows)

```powershell
powershell -NoProfile -File scripts/render_powerpoint.ps1 -InputPptx deck.pptx -OutputDirectory renders/run-01
```

Needs installed, licensed desktop PowerPoint and local COM access. Creates a new
output directory containing one 2560×1440 PNG per slide and a native vector PDF.
Pass a matching `-Width` and `-Height` for a non-16:9 deck. The report records
Office version and source SHA-256 before/after. It opens a read-only, untitled
copy and closes only that copy. Existing presentations remain open. It does not
save over, convert in place, or alter the source. A permission failure is an environment
limitation; do not call an SVG preview a finished PowerPoint render.

The primary PDF path is `ExportAsFixedFormat`. Some Office COM type libraries
cannot marshal its `PrintRange` parameter from PowerShell; in that case the
script records the error and uses native `SaveAs(ppSaveAsPDF)` on the untitled
copy. `pdf_method` records the method used. Inspect PDF fonts and raster content
when vector purity matters, especially after a fallback.

`-Svg` attempts PowerPoint's SVG filter and records support or failure per slide.
Some Office versions do not provide this filter. Native PDF is the primary
vector export here. Existing raster images in a slide do not become vectors.

Visually review every PNG at presentation scale and enlarged at dense areas.
Check titles, hierarchy, whitespace, formula placement, plot/table legibility,
line crossings, clipping, consistent color meaning, and content accuracy. Repair
only the relevant owning source, rerender changed pages and inspect neighbors.
A successful render and an XML pass do not establish aesthetic superiority.

## Explicit, non-raster PDF crop

```powershell
python scripts/enhanced_qa.py crop-pdf renders/run-01/presentation.pdf figures/method.pdf --page 2 --bbox 24 36 936 504 --report validation/crop.json
```

Coordinates use PDF points from the bottom left of an unrotated page. Select the
box after viewing the exported PDF; automatic ink-bound detection is deliberately
not claimed. The command rejects replacement of the source, an existing output,
invalid boxes, boxes beyond MediaBox, and rotated pages. It writes only the
selected page, changes page boxes without rasterizing, and verifies that its
decoded content stream remains identical. Cropping is not redaction; hidden
content can remain in the PDF. Confirm the new framing visually.

## Exact single-run text primitive

```powershell
python scripts/native_text_patch.py before.pptx after.pptx --slide 2 --shape-id 39 --old "一致性反馈" --new "对偶反馈" --report validation/text-patch.json
```

Use only when the old text is one complete, unique ordinary DrawingML text run
in the selected native shape. This does not handle a phrase split across runs,
math, tables or multiple objects. The original Edit Native roundtrip handles
those cases. Every other member payload and every other XML byte in the slide
is preserved; ZIP archive headers/compression can change. Run strict compare and
render the result. Longer text is not fitted automatically.
