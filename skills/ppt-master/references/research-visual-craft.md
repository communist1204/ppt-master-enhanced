# Research visual craft

Apply only to scientific diagrams, mechanism explanations and result pages.
Use the upstream native SVG-to-PPTX vocabulary and data/formula interfaces.

- Start with the page's question and its one useful conclusion. Let a method page
  show the path from input to output, a mechanism page explain one update, and a
  results page expose comparable evidence. Do not force all three into cards.
- Scientific architecture: differentiate state/data from operators by shape or
  labeling in addition to color. Keep connector direction and iteration boundaries
  unambiguous; use direct paths and restrained line weights. When a branch changes
  the algorithm state, connect it visibly to its source and destination instead of
  replacing it with an unattached note. Give independent feedback paths separate
  routing corridors so a crossing cannot imply a junction. Omit an arrow if the
  source establishes no dependency. Preserve meaningful branches and feedback loops.
- Mathematical mechanism: give the exact native equation enough vertical space;
  align its explanatory labels with the terms they describe. An auxiliary trajectory
  or schematic may aid intuition but label it as conceptual; never imply measured
  samples or a convergence guarantee absent from the source. Prefer one decisive
  equation to unreadable derivation wallpaper.
- Quantitative results: expose metric direction, dataset/condition and units.
  Align decimal values, keep comparison methods in a stable order, highlight only
  conclusions supported by the selected rows. A native table is better than a
  raster screenshot when the source values are available. A crop or selected row
  subset must be described. Do not invent missing measurements or significance.
- Qualitative results: use equal image scales/crops, consistent method labels and
  shared zoom boxes if available. Preserve the original result pixels. A source
  composite can be embedded honestly when separation would risk mislabeling it.
  Do not generate a new scientific result with an image model.
- Academic deck: use an opening with a clear problem and a closing with both the
  contribution and a real limit. Keep a common typographic/color system while
  allowing the evidence to change the composition. Increase the evidence area
  before shrinking its labels; footnotes must remain readable at the target use.

These are composition judgments, not a mandatory visual template. Small design
ideas are allowed and disclosed under `enhanced-delivery.md`. Source figures keep
their existing annotations; newly authored architecture and formulas should remain
native/editable. Treat final PPTX rendering as truth when it differs from SVG.
