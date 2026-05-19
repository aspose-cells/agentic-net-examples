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
- add-an-ole-object-for-a-powerpoint-slide-set-its-size-to-match-cell-g3-and-lock-it.cs
- load-a-workbook-add-an-ole-object-for-a-csv-file-and-configure-it-to-open-with-application.cs
- extract-the-embedded-ole-object-stream-from-cell-j7-write-it-to-a-temporary-file-and-close-workbook.cs
- extract-ole-object-data-decompress-if-necessary-and-save-the-original-file-format-to-a-specified-folder.cs
- extract-an-ole-object-rename-the-extracted-file-based-on-worksheet-name-and-save-to-output-directory.cs
- edit-an-ole-object-by-changing-its-source-file-path-to-a-new-word-document-and-update-properties.cs
- create-an-ole-object-for-a-word-document-embed-it-then-change-its-display-icon-to-an-image.cs
- group-a-picture-and-a-chart-lock-the-group-then-attempt-to-ungroup-to-test-lock-enforcement.cs
- create-a-shape-group-lock-it-then-attempt-to-move-an-inner-shape-to-test-lock-enforcement.cs
- ungroup-a-previously-created-shape-group-modify-each-individual-shapes-position-and-save-changes.cs
- batch-process-multiple-worksheets-by-inserting-a-picture-linked-to-cell-a1-in-each-sheet-and-saving-file.cs
- batch-insert-pictures-from-a-csv-list-of-image-urls-linking-each-picture-to-its-corresponding-cell-reference.cs
- load-an-existing-workbook-and-access-its-worksheet-collection-for-manipulation.cs
- insert-a-wav-file-as-an-embedded-ole-object-at-specified-cell-coordinates.cs
- set-the-oleobjectname-property-to-a-unique-identifier-after-insertion.cs
- configure-the-oleobjectwidth-and-height-to-fit-within-the-target-cell-range.cs
- set-the-oleobjectlockaspectratio-flag-to-preserve-original-proportions-during-resizing.cs
- assign-a-custom-display-label-to-the-ole-object-using-the-label-property.cs
- retrieve-an-oleobject-by-name-from-the-worksheets-oleobjects-collection-for-modification.cs
- change-the-position-of-a-retrieved-oleobject-using-top-and-left-offset-values.cs
- update-the-display-label-of-a-linked-ole-object-to-a-descriptive-string.cs
- extract-an-embedded-ole-object-to-a-designated-output-folder-on-disk.cs
- verify-that-the-extracted-files-extension-matches-the-original-format-for-integrity.cs
- read-the-classid-clsid-property-of-an-embedded-ole-object-for-auditing.cs
- assign-a-new-guid-to-the-oleobjectclassid-to-change-its-associated-application.cs
- iterate-through-all-oleobjects-in-a-workbook-and-extract-each-to-a-folder.cs
- use-the-oleobjectislinked-property-to-identify-linked-ole-objects-before-processing.cs
- refresh-linked-ole-objects-by-invoking-excel-automation-through-the-updatelink-method.cs
- save-the-workbook-after-refreshing-linked-ole-objects-to-apply-the-updates.cs
- export-a-worksheet-containing-ole-objects-to-pdf-while-preserving-object-placeholders.cs
- configure-the-workbook-to-run-in-licensed-mode-to-suppress-evaluation-watermarks.cs
- log-the-original-file-name-and-size-of-each-extracted-ole-object-for-traceability.cs
- create-a-batch-process-that-adds-the-same-wav-ole-object-to-every-worksheet.cs
- implement-error-handling-to-catch-exceptions-when-extracting-corrupted-or-unsupported-ole-objects.cs
- generate-a-report-listing-each-ole-objects-name-label-class-identifier-and-file-size.cs
- use-oleobjectalternativetext-property-to-store-custom-metadata-for-later-retrieval.cs
- clone-an-existing-ole-object-and-place-the-copy-on-a-different-worksheet.cs
- remove-an-unwanted-ole-object-from-a-worksheet-based-on-its-label-content.cs
- add-a-hyperlink-to-an-ole-object-that-opens-the-original-source-file-when-clicked.cs
- set-the-oleobjectvisible-property-to-false-for-background-objects-that-should-not-appear.cs
- load-the-workbook-and-select-the-target-worksheet-before-performing-any-shape-operations.cs
- add-a-new-textbox-to-the-worksheet-at-the-specified-cell-coordinates.cs
- assign-a-unique-name-to-the-textbox-for-later-identification-and-manipulation.cs
- retrieve-the-textbox-by-its-assigned-name-to-modify-its-properties-programmatically.cs
- set-the-displayed-text-of-the-textbox-using-a-provided-string-value.cs
- append-additional-text-to-the-existing-content-of-the-textbox-programmatically.cs
- apply-left-alignment-to-the-entire-text-within-the-textbox.cs
- apply-center-alignment-to-selected-characters-inside-the-textbox-using-rich-text-formatting.cs
- apply-right-alignment-to-specific-characters-within-the-textbox-for-emphasis.cs
- change-the-font-size-of-partial-text-inside-the-textbox-to-highlight-important-words.cs
- apply-bold-and-italic-styles-to-selected-text-fragments-within-the-textbox.cs
- change-the-font-color-of-partial-text-inside-the-textbox-to-a-custom-rgb-value.cs
- adjust-the-internal-margins-of-the-textbox-to-control-padding-around-the-text.cs
- enable-text-wrapping-inside-the-textbox-so-long-sentences-break-onto-multiple-lines.cs
- configure-autofit-for-the-textbox-to-resize-automatically-based-on-its-content-length.cs
- rotate-the-textbox-by-a-specified-angle-and-verify-its-new-orientation-on-the-worksheet.cs
- set-the-textbox-background-fill-to-a-solid-color-with-optional-transparency.cs
- apply-a-gradient-fill-to-the-textbox-using-two-custom-colors-and-a-defined-angle.cs
- add-a-hyperlink-to-the-textbox-that-opens-a-web-page-when-the-shape-is-clicked.cs
- lock-the-textbox-to-prevent-users-from-moving-or-resizing-it-in-the-excel-ui.cs
- unlock-a-previously-locked-textbox-to-allow-editing-of-its-position-and-size.cs
- set-alternative-text-for-the-textbox-to-improve-accessibility-for-screen-readers.cs
- export-the-textbox-as-a-png-image-with-a-transparent-background-for-external-use.cs
- delete-the-textbox-from-the-worksheet-when-it-is-no-longer-required.cs
- copy-the-textbox-to-another-worksheet-while-preserving-its-size-and-text-attributes.cs
