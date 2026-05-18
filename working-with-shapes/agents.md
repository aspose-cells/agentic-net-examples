# Working With Shapes Examples

This folder contains **Aspose.Cells for .NET** code examples related to:

Working With Shapes


## Purpose

These examples demonstrate common **Aspose.Cells APIs** used when working with:

- Workbooks
- Worksheets
- Cells
- Formulas
- Charts
- Data operations


## Example Files

Each `.cs` file demonstrates a specific task related to **Working With Shapes**.

Example:

create-a-workbook.cs


## Required Namespaces

Most examples will require:

using Aspose.Cells;


## Common Pattern

Typical Aspose.Cells workflow:

Workbook workbook = new Workbook();

Worksheet sheet = workbook.Worksheets[0];

Cells cells = sheet.Cells;


## Output

Examples may generate:

- XLSX files
- PDF files
- CSV files
- Images

Output files are written to the working directory.
- insert-a-rectangle-shape-of-a-specific-type-at-the-target-cell-location-in-the-worksheet.cs
- add-a-shape-with-custom-width-and-height-dimensions-anchored-to-a-designated-cell-range.cs
- place-a-shape-and-set-its-rotation-angle-to-ninety-degrees-for-diagonal-orientation.cs
- insert-a-shape-and-configure-a-linear-gradient-fill-transitioning-between-two-chosen-colors.cs
- insert-a-shape-and-attach-a-hyperlink-directing-users-to-an-external-website-when-clicked.cs
- create-a-shape-and-embed-a-comment-that-appears-as-a-tooltip-on-mouse-hover.cs
- add-a-shape-and-set-alternative-text-describing-its-purpose-for-accessibility-compliance.cs
- insert-a-shape-and-assign-a-unique-name-property-to-reference-it-in-later-code.cs
- place-a-shape-and-adjust-its-z-order-to-bring-it-forward-above-other-objects.cs
- add-a-shape-and-enable-its-locked-property-to-prevent-accidental-modifications.cs
- insert-a-shape-and-toggle-its-visibility-property-to-hide-it-from-the-worksheet-view.cs
- create-a-shape-and-apply-conditional-formatting-that-changes-its-fill-based-on-cell-values.cs
- add-a-shape-that-incorporates-data-validation-rules-to-restrict-user-input-within-defined-limits.cs
- insert-a-picture-from-a-local-file-path-and-position-it-at-a-specified-cell-coordinate.cs
- add-a-picture-loaded-from-a-web-url-and-embed-it-directly-into-the-worksheet.cs
- insert-a-picture-and-scale-its-dimensions-proportionally-to-fit-within-a-target-cell-range.cs
- add-a-picture-and-apply-cropping-parameters-to-display-only-the-central-portion-of-the-image.cs
- insert-a-picture-and-set-its-transparency-level-to-achieve-a-semitransparent-visual-effect.cs
- add-a-picture-and-rotate-it-ninety-degrees-clockwise-to-align-with-column-orientation.cs
- insert-a-picture-and-attach-a-hyperlink-that-opens-a-document-when-the-image-is-clicked.cs
- add-a-picture-and-provide-alternative-text-describing-its-content-for-screen-reader-accessibility.cs
- insert-a-picture-and-assign-a-descriptive-name-property-to-facilitate-later-retrieval-via-api.cs
- add-a-picture-and-enable-aspect-ratio-lock-to-maintain-proportional-dimensions-during-resizing.cs
- insert-a-picture-and-apply-background-removal-to-isolate-the-foreground-subject-from-its-backdrop.cs
- add-a-picture-and-overlay-a-semitransparent-watermark-to-protect-intellectual-property-rights.cs
- create-a-camera-shape-for-a-specified-range-to-capture-its-visual-representation-as-an-image.cs
- configure-a-camera-shape-to-capture-the-target-range-with-defined-resolution-and-image-format-settings.cs
- refresh-a-camera-shape-programmatically-using-the-api-method-to-ensure-current-range-content-is-shown.cs
- load-a-workbook-file-insert-a-picture-linked-to-cell-a1-and-save-as-xlsx.cs
- load-a-workbook-from-stream-attach-a-linked-picture-from-web-url-to-cell-g10-then-refresh-links.cs
- insert-a-linked-picture-from-an-image-url-set-islinked-true-and-ensure-image-data-is-not-embedded.cs
- insert-a-linked-picture-from-a-secure-intranet-url-configure-authentication-headers-and-handle-access-denied-errors.cs
- refresh-all-linked-pictures-in-the-workbook-after-updating-source-images-on-the-web-server.cs
- refresh-linked-pictures-after-modifying-source-images-on-a-local-file-system-ensuring-updated-visuals-appear.cs
- refresh-linked-pictures-in-parallel-threads-to-improve-performance-when-updating-dozens-of-external-images.cs
- insert-a-linked-picture-then-programmatically-change-its-source-url-to-a-new-image-and-refresh-it.cs
- insert-a-picture-linked-to-cell-h12-then-programmatically-change-its-upperleftcell-to-i13-and-verify-position.cs
- add-a-picture-linked-to-cell-c9-then-retrieve-its-absolute-x-and-y-coordinates-for-logging.cs
- configure-a-picture-to-move-and-size-with-its-linked-cell-then-verify-resizing-cell-adjusts-picture.cs
- create-a-picture-object-set-upperleftcell-to-d5-enable-move-with-cell-and-add-to-worksheet.cs
- insert-a-picture-set-its-height-to-200-points-and-maintain-aspect-ratio-automatically.cs
- insert-a-picture-set-its-z-order-to-bring-it-forward-behind-only-the-chart-shape-and-save.cs
- send-a-chart-shape-backward-so-that-a-linked-picture-appears-above-it-in-the-worksheet.cs
- retrieve-the-absolute-position-of-a-shape-anchored-at-cell-f5-and-log-its-x-and-y-coordinates.cs
- link-a-shapes-text-to-cell-k3-so-the-shape-updates-automatically-when-the-cell-value-changes.cs
- apply-a-conditional-icon-set-for-range-b2b15-enable-showcellvalue-and-define-custom-icons-for-each-condition.cs
- apply-an-icon-set-condition-to-cells-c3c12-replace-default-icons-with-text-labels-and-save-workbook.cs
- apply-a-conditional-icon-set-with-custom-green-yellow-and-red-icons-to-a-range-of-sales-figures.cs
- create-a-conditional-formatting-rule-using-an-icon-set-to-display-arrows-based-on-numeric-thresholds.cs
- apply-a-conditional-formatting-rule-using-an-icon-set-to-display-custom-emoji-icons-based-on-text-values.cs
- load-a-workbook-add-a-conditional-icon-set-with-three-custom-png-icons-and-save-as-xlsx.cs
- load-a-workbook-add-a-conditional-icon-set-that-shows-cell-values-and-save-file-as-xlsm-macro-enabled.cs
- create-an-ole-object-from-a-pdf-file-embed-it-into-worksheet-at-cell-h4-and-set-size.cs
- load-a-workbook-add-an-ole-object-for-a-pdf-set-its-display-mode-to-icon-and-save.cs
- create-an-ole-object-for-a-pdf-set-its-icon-caption-to-report-and-embed-it-in-sheet.cs
- insert-an-ole-object-representing-a-visio-diagram-set-its-size-to-match-cell-dimensions-and-lock-it.cs
