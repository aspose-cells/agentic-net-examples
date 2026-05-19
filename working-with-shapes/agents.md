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
- move-the-textbox-to-a-different-cell-location-by-updating-its-anchor-coordinates.cs
- ungroup-previously-grouped-shapes-ensuring-each-retains-its-original-formatting.cs
- set-the-zorder-of-the-textbox-to-bring-it-to-the-front-of-overlapping-objects.cs
- load-a-workbook-locate-a-textbox-named-headerbox-and-replace-tag_1-with-dynamic-title.cs
- validate-that-all-textbox-shapes-contain-required-tags-before-performing-batch-replacement-to-avoid-missing-data-errors.cs
- replace-multiple-placeholder-tags-tag_a-tag_b-and-tag_c-within-a-textbox-using-a-dictionary-mapping.cs
- implement-error-handling-to-catch-exceptions-when-a-specified-textbox-name-does-not-exist-in-the-worksheet.cs
- create-a-new-textbox-on-sheet1-set-its-width-to-200-points-and-position-it-at-cell-b2.cs
- create-a-multiline-textbox-and-set-individual-line-alignments-to-left-center-and-right-respectively.cs
- apply-left-alignment-to-the-first-line-and-center-alignment-to-the-second-line-of-a-textbox.cs
- load-multiple-svg-icons-from-a-folder-and-place-each-icon-into-successive-rows-starting-from-row-ten.cs
- batch-insert-a-company-logo-svg-at-the-topright-corner-of-each-worksheet-in-the-workbook.cs
- load-a-workbook-from-a-url-replace-tags-in-all-textboxes-and-save-the-file-to-cloud-storage.cs
- iterate-over-all-textbox-shapes-extract-their-inner-text-and-write-the-collected-strings-to-a-csv-file.cs
- iterate-all-shapes-on-a-worksheet-listing-each-shapes-name-type-and-absolute-coordinates.cs
- retrieve-the-absolute-x-and-y-coordinates-of-shape-logo-and-log-them-for-debugging.cs
- send-a-shape-named-chartoverlay-to-the-front-layer-to-ensure-it-appears-above-all-objects.cs
- move-a-shape-called-watermark-to-the-back-layer-so-underlying-cells-remain-visible.cs
- link-a-shape-commentbox-to-cell-g12-so-the-shape-moves-when-the-cell-shifts.cs
- configure-a-shape-to-follow-cell-resizing-by-linking-it-to-cell-h5-and-enabling-movewithcells-option.cs
- access-connection-points-of-shape-flowconnector-and-attach-them-to-cells-a1-and-b2-for-dynamic-linking.cs
- group-shapes-icon1-icon2-and-icon3-into-a-single-group-named-iconset-for-collective-manipulation.cs
- lock-the-shape-signature-to-prevent-accidental-resizing-or-repositioning-during-workbook-editing.cs
- unlock-the-shape-signature-after-confirmation-to-allow-modifications-to-its-size-and-position.cs
- duplicate-an-existing-shape-rename-the-copy-and-offset-its-position-by-ten-points-to-the-right.cs
- export-the-formula-in-cell-c5-to-latex-using-tolatex-and-embed-the-result-in-an-html-paragraph.cs
- convert-the-formula-in-cell-d10-to-mathml-via-tomathml-and-store-the-markup-in-a-string-variable.cs
- export-multiple-worksheet-formulas-to-latex-and-concatenate-the-results-into-a-single-document.cs
- generate-mathml-for-each-formula-in-column-h-and-store-the-markup-in-adjacent-cells.cs
- create-a-shape-that-displays-a-live-formula-result-by-linking-its-text-to-the-calculation-cell.cs
- apply-conditional-formatting-to-a-shapes-fill-color-based-on-a-linked-cells-value-exceeding-a-threshold.cs
- resize-a-shape-proportionally-to-match-the-dimensions-of-a-target-cell-range-while-maintaining-aspect-ratio.cs
- programmatically-remove-all-shapes-from-a-worksheet-before-exporting-the-data-to-a-csv-format-for-clean-output.cs
- export-all-shapes-xml-definitions-to-a-separate-file-for-external-analysis-and-version-control-tracking.cs
- export-worksheet-shapes-as-separate-svg-files-preserving-visual-properties-for-reuse-in-web-pages.cs
- validate-inserted-svg-files-to-ensure-they-do-not-contain-unsupported-elements-that-could-cause-rendering-errors.cs
- implement-a-batch-operation-that-replaces-date-tags-in-all-textboxes-with-the-current-system-date-formatted-yyyymmdd.cs
- detect-shapes-that-exceed-worksheet-boundaries-and-automatically-reposition-them-within-visible-limits-to-avoid-clipping.cs
- change-the-line-style-of-shape-divider-to-dashed-and-set-its-thickness-to-two-points.cs
- create-a-shape-containing-a-hyperlink-to-another-worksheet-within-the-same-workbook-for-quick-navigation.cs
- add-a-textbox-control-to-the-first-worksheet-and-set-its-initial-text.cs
- retrieve-all-textbox-controls-from-a-worksheet-and-output-their-positions.cs
- change-the-text-of-a-specific-textbox-and-save-the-workbook-as-xlsx.cs
- add-a-checkbox-control-to-a-worksheet-and-link-its-state-to-cell-b2.cs
- update-the-linked-cell-reference-of-an-existing-checkbox-to-reflect-layout-changes.cs
- create-a-combobox-control-populate-it-with-five-items-and-define-its-dropdown-width.cs
- read-the-currently-selected-value-from-an-activex-combobox-on-the-second-sheet.cs
- update-an-activex-combobox-value-to-a-custom-string-and-verify-the-change-programmatically.cs
- add-a-listbox-control-enable-multiselection-mode-and-populate-it-with-ten-entries.cs
- set-the-selected-indices-of-a-listbox-based-on-values-from-column-a.cs
- insert-a-button-control-assign-a-macro-name-and-position-it-at-cell-d5.cs
- configure-the-button-to-trigger-a-vba-script-that-highlights-the-active-row.cs
- resize-the-line-to-span-cells-b2-through-e2-and-adjust-its-thickness.cs
- create-a-rectangle-shape-fill-it-with-light-blue-and-place-it-behind-chart-objects.cs
- group-multiple-rectangle-shapes-into-a-single-container-for-collective-movement.cs
- insert-a-generic-activex-togglebutton-using-shapecollectionaddactivexcontrol-and-set-its-default-state.cs
- access-the-newly-added-activex-control-via-shapeactivexcontrol-and-cast-it-to-its-specific-type.cs
- set-the-size-and-position-of-an-activex-control-after-insertion-using-the-shape-objects-properties.cs
- remove-an-existing-activex-control-from-a-worksheet-and-verify-its-absence-in-the-shapes-collection.cs
- batch-add-textbox-controls-to-each-worksheet-assigning-unique-identifiers.cs
- iterate-through-all-worksheets-and-delete-any-checkbox-controls-linked-to-cells-containing-zero.cs
- export-the-properties-of-all-shape-controls-to-a-json-file-for-external-analysis.cs
- import-shape-configuration-from-a-json-file-and-recreate-the-controls-on-a-target-worksheet.cs
- apply-conditional-formatting-to-a-textbox-based-on-the-numeric-value-of-its-linked-cell.cs
- synchronize-the-checked-state-of-multiple-checkbox-controls-with-corresponding-cells-in-a-data-table.cs
- validate-that-each-activex-control-has-a-unique-name-property-before-saving-the-workbook.cs
- generate-a-report-listing-all-shape-types-present-in-a-workbook-along-with-their-cell-coordinates.cs
- create-a-new-workbook-and-insert-a-wordart-watermark-on-the-first-worksheet.cs
- adjust-the-wordart-shapes-top-left-width-and-height-to-cover-the-entire-worksheet.cs
- send-the-inserted-wordart-shape-to-the-back-of-the-sheet-to-act-as-background-watermark.cs
- add-a-second-wordart-shape-with-builtin-style-simple-fill-and-custom-font-size.cs
- apply-the-preset-wordart-style-wave-to-the-first-shapes-text-using-fontsetting.cs
- use-fontsettingcollection-to-apply-the-bold-wave-preset-style-to-all-worksheet-wordart-shapes.cs
- group-the-two-wordart-shapes-together-and-lock-the-group-to-maintain-layout-integrity.cs
- iterate-through-all-worksheets-and-add-a-semitransparent-wordart-watermark-to-each-sheet.cs
- save-the-workbook-containing-wordart-watermarks-as-an-xlsx-file-preserving-all-shape-attributes.cs
- export-the-workbook-with-wordart-watermarks-to-pdf-format-ensuring-watermark-visibility-in-output.cs
- load-an-existing-excel-file-replace-its-header-wordart-with-a-new-style-and-save.cs
- create-a-wordart-shape-set-rotation-to-45-degrees-and-position-it-diagonally-across-the-sheet.cs
- configure-the-wordart-shapes-fill-to-a-gradient-transitioning-from-blue-to-transparent.cs
- apply-builtin-wordart-style-curved-text-to-a-shape-and-adjust-curvature-via-adjustment-values.cs
- duplicate-the-wordart-shape-move-the-copy-to-a-different-cell-range-and-change-its-text.cs
- determine-zorder-of-all-shapes-programmatically-and-bring-the-most-important-wordart-to-front.cs
- create-a-routine-that-removes-all-existing-wordart-shapes-before-adding-new-watermarks.cs
- validate-each-worksheet-contains-exactly-one-locked-wordart-watermark-and-report-any-discrepancies.cs
- implement-error-handling-to-catch-exceptions-when-adding-wordart-to-a-protected-worksheet.cs
- apply-preset-wordart-style-bold-wave-to-shapes-only-on-worksheets-whose-names-start-with-report.cs
- create-a-batch-process-that-reads-multiple-excel-files-from-a-directory-and-adds-consistent-watermark.cs
- implement-logging-to-record-timestamp-and-worksheet-name-each-time-a-wordart-shape-is-added.cs
- create-a-custom-function-that-returns-true-if-a-given-shape-is-a-locked-wordart-watermark.cs
- use-linq-to-filter-shapecollection-for-only-wordart-shapes-before-applying-batch-style-changes.cs
- set-wordart-shapes-text-alignment-to-center-horizontally-and-vertically-within-the-shape-bounds.cs
- adjust-wordart-shapes-transparency-to-thirty-percent-to-create-a-subtle-background-effect.cs
- create-a-reusable-method-that-adds-a-wordart-watermark-with-specified-text-style-and-lock-settings.cs
- apply-builtin-wordart-style-shadowed-to-a-shape-and-modify-its-shadow-offset-manually.cs
- create-a-wordart-shape-set-flip-horizontally-property-and-verify-visual-reversal.cs
- add-a-wordart-shape-to-a-specific-cell-range-and-anchor-it-so-it-moves-with-cells.cs
- programmatically-unlock-a-previously-locked-wordart-watermark-to-allow-user-modifications-in-the-workbook.cs
- create-a-wordart-shape-assign-a-custom-font-family-and-set-the-font-size-to-fortyeight-points.cs
- use-shapecollectionaddwordart-overload-that-accepts-a-style-name-to-directly-apply-a-builtin-style.cs
- iterate-over-all-shapes-identify-wordart-objects-and-change-their-text-color-to-dark-blue.cs
- create-a-wordart-shape-set-adjustment-values-to-create-custom-curvature-and-save-the-workbook.cs
- group-three-wordart-shapes-lock-the-group-and-export-the-worksheet-as-an-xlsx-file.cs
- transfer-a-chart-from-sheet1-to-sheet3-using-shapesaddcopy-while-maintaining-data-source-references.cs
- duplicate-form-controls-from-one-worksheet-to-another-ensuring-linked-cell-addresses-remain-unchanged.cs
- group-three-selected-shapes-into-a-composite-object-and-assign-a-descriptive-name-for-reference.cs
- ungroup-a-previously-grouped-shape-collection-to-modify-individual-components-formatting-and-positioning.cs
- lock-a-specific-picture-shape-to-prevent-editing-during-worksheet-protection-and-verify-its-locked-property.cs
- unlock-all-shapes-on-a-worksheet-by-iterating-through-shapecollection-and-clearing-each-shapes-locked-flag.cs
- apply-a-soft-shadow-effect-with-custom-offset-and-color-to-a-chart-shape-for-visual-depth.cs
- add-a-glowing-outer-border-to-a-picture-shape-using-specified-glow-radius-and-color-value.cs
- implement-a-reflection-effect-on-a-shape-with-defined-transparency-and-size-parameters-to-simulate-glass.cs
- configure-3d-rotation-and-bevel-properties-on-a-shape-to-create-realistic-threedimensional-appearance.cs
- adjust-a-shapes-adjustment-values-to-modify-its-geometry-such-as-changing-a-stars-point-count.cs
- retrieve-the-absolute-topleft-coordinates-of-a-shape-to-align-it-precisely-with-a-target-cell-range.cs
- obtain-all-connection-points-of-a-shape-to-calculate-attachment-positions-for-linked-arrows-and-connectors.cs
- send-a-specific-shape-to-the-front-of-the-zorder-stack-to-ensure-it-overlays-other-objects.cs
- move-a-shape-to-the-back-of-the-zorder-hierarchy-so-underlying-cells-become-visible-through-it.cs
- assign-a-cell-range-anchor-to-a-shape-so-it-moves-dynamically-when-rows-or-columns-are-inserted.cs
- read-the-current-shadow-color-of-a-shape-and-output-its-rgb-values-for-logging-purposes.cs
- extract-the-glow-color-from-a-shape-and-compare-it-against-a-predefined-palette-for-consistency-checks.cs
- batch-copy-all-pictures-from-multiple-worksheets-into-a-summary-sheet-preserving-each-pictures-original-size.cs
- create-a-macrolike-routine-that-groups-shapes-by-type-then-locks-each-group-to-prevent-accidental-edits.cs
- develop-a-validation-script-that-ensures-no-shape-exceeds-a-specified-width-and-height-threshold-across-the-workbook.cs
- write-code-to-clone-a-chart-shape-modify-its-data-series-and-place-it-on-another-worksheet.cs
- design-a-routine-that-extracts-all-shape-hyperlinks-and-generates-a-report-listing-source-cells-and-target-urls.cs
- programmatically-set-the-zorder-of-a-shape-based-on-its-importance-level-defined-in-a-custom-metadata-field.cs
- develop-a-process-to-ungroup-nested-shape-groups-recursively-until-only-individual-shapes-remain-on-the-sheet.cs
- implement-error-handling-that-logs-a-warning-when-attempting-to-lock-a-shape-already-marked-as-locked.cs
- create-a-script-that-applies-a-uniform-glow-effect-to-every-shape-whose-name-contains-important.cs
- develop-a-method-to-retrieve-and-display-absolute-positions-of-all-shapes-on-a-worksheet-for-analysis.cs
- implement-a-batch-operation-that-locks-all-shapes-on-every-worksheet-before-exporting-the-workbook-to-pdf.cs
- write-code-to-set-a-pictures-anchor-to-a-merged-cell-range-ensuring-it-moves-with-merged-cells.cs
- create-a-routine-that-reads-each-shapes-shadow-color-and-adjusts-it-to-match-the-workbooks-theme-palette.cs
- develop-a-feature-that-automatically-groups-shapes-added-within-a-specific-time-window-into-a-single-collection.cs
- write-a-function-that-calculates-the-distance-between-two-shapes-using-their-absolute-positions-for-collision-detection.cs
- implement-a-process-to-copy-a-shapes-3d-format-settings-from-one-shape-to-another-across-worksheets.cs
- create-a-validation-rule-that-ensures-every-shape-linked-to-a-cell-references-a-valid-nonempty-cell-address.cs
- develop-a-script-that-iterates-through-all-shapes-unlocking-those-with-a-specific-tag-attribute-for-editing.cs
- write-code-to-apply-a-custom-reflection-gradient-to-a-picture-shape-adjusting-both-opacity-and-size-parameters.cs
- implement-a-batch-copy-operation-that-transfers-all-drawing-objects-from-a-template-worksheet-to-newly-created-worksheets.cs
- create-a-utility-that-lists-each-shapes-zorder-index-and-sorts-them-alphabetically-by-name-for-reporting.cs
- write-a-routine-that-copies-a-chart-shape-changes-its-title-and-places-it-on-a-summary-worksheet.cs
- implement-a-feature-that-automatically-unlocks-all-shapes-when-the-workbook-is-opened-in-edit-mode.cs
- create-a-script-that-applies-a-uniform-shadow-offset-to-every-shape-ensuring-consistent-visual-alignment.cs
- write-code-that-links-each-shape-to-its-corresponding-data-cell-enabling-dynamic-updates-when-cell-values-change.cs
- load-an-excel-workbook-from-a-file-and-access-the-first-worksheet.cs
- apply-a-50-transparent-reflection-to-the-selected-shape-using-shapereflection.cs
- set-reflection-size-to-30-points-and-blur-radius-to-5-points-on-the-shape.cs
- configure-shadow-effect-with-45-degree-angle-10-point-distance-and-dark-gray-color.cs
- apply-an-outer-shadow-preset-and-set-its-transparency-to-40-percent-for-the-shape.cs
- define-a-blue-glow-effect-with-size-eight-points-on-all-shapes-in-the-worksheet.cs
- read-the-current-glow-color-of-a-specific-shape-and-log-it-to-the-console.cs
- set-threedimensional-material-to-plastic-and-apply-lighting-preset-soft-with-seventy-percent-intensity.cs
- rotate-the-threedimensional-shape-by-thirty-degrees-around-the-x-axis.cs
- tile-a-picture-file-as-texture-inside-a-rectangle-shape-while-preserving-aspect-ratio.cs
- send-the-shape-to-the-front-layer-to-adjust-zorder-positioning.cs
- retrieve-all-connection-points-of-a-chart-shape-and-export-coordinates-to-a-csv-file.cs
- update-a-shapes-connection-point-offset-to-1020-and-persist-the-changes.cs
- apply-a-custom-reflection-preset-with-size-forty-blur-three-and-distance-six-points.cs
- set-shadow-color-using-argb-values-25512800-and-transparency-to-thirty-percent.cs
- create-a-gradient-glow-transitioning-from-red-to-yellow-with-size-five-points.cs
- apply-threedimensional-lighting-direction-top-and-set-specular-power-to-eighty-for-the-shape.cs
- clone-a-shape-modify-its-reflection-and-place-the-clone-beside-the-original-shape.cs
- remove-the-shadow-effect-from-a-shape-while-preserving-its-reflection-and-glow.cs
- set-threedimensional-rotation-angles-x20-y45-z10-for-a-chart-and-save.cs
- set-the-reflection-distance-to-twelve-points-and-verify-visual-change-using-automated-screenshot-comparison.cs
- disable-all-shape-reflections-while-preserving-other-effects-and-save-the-workbook.cs
- copy-reflection-settings-from-one-shape-to-another-within-the-same-worksheet.cs
- toggle-the-visibility-of-the-glow-effect-on-a-specific-shape-without-altering-other-properties.cs
- remove-all-threedimensional-formatting-from-shapes-while-keeping-twodimensional-properties-intact.cs
- load-an-excel-workbook-from-a-file-path-and-access-the-first-worksheet.cs
- retrieve-a-shape-by-its-name-from-the-worksheet-and-verify-its-existence.cs
- assign-a-png-image-stream-to-the-shapes-filltexturefillimage-property.cs
- enable-tiling-on-the-shape-by-setting-filltexturefillistiling-to-true.cs
- save-the-modified-workbook-to-a-new-file-preserving-original-formatting-and-textures.cs
- iterate-through-all-shapes-in-the-worksheet-and-apply-tiled-texture-to-each.cs
- load-an-image-from-a-url-and-use-it-as-texture-for-a-specific-shape.cs
- create-a-new-text-box-shape-and-set-its-text-using-textbodytext-property.cs
- retrieve-the-fontsetting-object-from-a-shapes-textbody-to-modify-text-appearance.cs
- set-character-spacing-to-5-points-on-the-shapes-text-via-textoptionscharacterspacing.cs
- apply-a-preset-shadow-of-type-outershadow-to-the-shapes-text-via-textoptionsshadowpresettype.cs
- change-the-shadow-preset-to-perspectivediagonalbottomleft-and-observe-visual-difference.cs
- adjust-line-spacing-to-15-lines-within-the-shapes-text-using-textoptionslinespacing.cs
- rotate-the-shapes-text-by-45-degrees-using-textoptionsrotationangle-property.cs
- validate-that-the-shapes-texture-image-file-exists-before-assigning-it-to-avoid-errors.cs
- implement-error-handling-for-missing-image-files-when-setting-shape-texture.cs
- batch-process-multiple-workbooks-in-a-folder-adding-tiled-textures-to-a-designated-shape.cs
- measure-performance-impact-of-enabling-texture-tiling-on-large-worksheets-with-many-shapes.cs
- compare-visual-quality-between-tiled-texture-and-stretched-texture-by-toggling-istiling.cs
- use-presetshadowtypeinnershadow-to-create-a-subtle-inset-effect-on-shape-text.cs
- apply-a-custom-character-spacing-of-2-points-to-compress-text-within-a-narrow-shape.cs
- programmatically-duplicate-a-shape-and-assign-a-different-texture-image-to-the-copy.cs
- retrieve-the-current-shadow-preset-of-a-shapes-text-and-log-it-for-debugging.cs
- validate-that-character-spacing-values-remain-within-the-allowed-range-of-10-to-10-points.cs
- create-a-unit-test-that-verifies-texture-tiling-is-applied-when-istiling-is-true.cs
- develop-a-helper-method-that-returns-a-fontsetting-configured-with-a-specific-shadow-preset-and-spacing.cs
- iterate-over-all-text-boxes-in-a-worksheet-and-set-their-character-spacing-uniformly.cs
- use-a-trycatch-block-to-handle-exceptions-when-loading-an-invalid-image-for-texture.cs
- create-a-macrolike-routine-that-toggles-texture-tiling-on-a-shape-each-execution.cs
- apply-a-preset-shadow-of-type-outerbottom-to-shape-text-and-capture-a-screenshot.cs
- read-shape-properties-into-a-json-object-for-external-reporting-of-texture-and-text-settings.cs
- write-a-script-that-loads-a-workbook-from-a-url-applies-texture-tiling-and-saves-to-cloud-storage.cs
- implement-a-feature-that-disables-texture-tiling-when-the-shapes-area-is-smaller-than-the-image-dimensions.cs
