# Rows and Columns Examples

This folder contains **Aspose.Cells for .NET** code examples related to:

Rows and Columns


## Purpose

These examples demonstrate common **Aspose.Cells APIs** used when working with:

- Workbooks
- Worksheets
- Cells
- Formulas
- Charts
- Data operations


## Example Files

Each `.cs` file demonstrates a specific task related to **Rows and Columns**.

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
- load-a-workbook-with-loadoptionsautofitteroptionsonlyauto-set-to-true-to-automatically-adjust-all-row-heights.cs
- set-the-height-of-a-specific-row-eg-row-5-to-a-defined-point-value-using-cellssetrowheight.cs
- apply-a-uniform-row-height-to-all-rows-in-a-worksheet-by-assigning-cellsstandardheight.cs
- set-the-width-of-a-specific-column-eg-column-3-using-cellssetcolumnwidth.cs
- set-the-width-of-a-specific-column-in-pixels-using-cellssetcolumnwidthpixel.cs
- apply-a-uniform-column-width-to-all-columns-by-assigning-cellsstandardwidth.cs
- autofit-a-single-row-based-on-its-content-using-worksheetautofitrow.cs
- autofit-a-range-of-rows-eg-rows-1520-using-worksheetautofitrows.cs
- autofit-a-single-column-based-on-its-content-using-worksheetautofitcolumn.cs
- autofit-a-range-of-columns-eg-columns-cf-using-worksheetautofitcolumns.cs
- autofit-rows-that-contain-merged-cells-by-configuring-autofitteroptionsautofitmergedcellstype-and-passing-it-to-worksheetautofitrows.cs
- combine-setrowheight-with-autofitrow-to-set-a-minimum-height-before-autofitting-a-row.cs
- after-autofitting-a-column-finetune-its-width-with-setcolumnwidthpixel-for-precise-pixel-control.cs
- iterate-through-each-worksheet-in-a-workbook-and-set-standardheight-to-enforce-a-consistent-row-height.cs
- batch-process-multiple-worksheets-to-apply-standardwidth-for-consistent-column-sizing.cs
- use-autofitrows-with-custom-autofitteroptions-to-ignore-merged-cells-while-adjusting-a-specific-row-range.cs
- load-a-workbook-modify-a-cell-value-then-autofit-the-affected-row-to-reflect-the-change.cs
- set-row-heights-for-a-series-of-rows-using-a-loop-that-calls-setrowheight-with-incremental-values.cs
- apply-setcolumnwidthpixel-to-a-group-of-columns-after-autofitting-them-to-achieve-precise-pixel-alignment.cs
- create-a-new-workbook-add-data-to-a-column-and-autofit-that-column-to-accommodate-the-longest-entry.cs
- use-worksheetautofitrows-overload-with-startrow-and-endrow-to-adjust-a-block-of-rows.cs
- enable-onlyauto-loading-then-immediately-save-the-workbook-as-pdf-to-produce-a-document-with-prefitted-rows.cs
- after-loading-a-workbook-set-standardheight-and-then-autofit-rows-that-contain-formulas.cs
- adjust-column-width-by-setting-standardwidth-before-populating-data-to-establish-a-base-width.cs
- autofit-rows-that-contain-merged-cells-by-specifying-autofitmergedcellstype-in-autofitteroptions-during-processing.cs
- programmatically-set-a-row-height-then-autofit-the-next-row-based-on-its-content.cs
- apply-setcolumnwidth-to-a-column-then-autofit-an-adjacent-column-for-comparison.cs
- load-a-workbook-with-onlyauto-enabled-then-iterate-through-each-sheet-to-verify-row-heights-are-adjusted.cs
- use-autofitcolumn-on-a-column-after-inserting-multiline-text-to-ensure-full-visibility.cs
- set-standardwidth-then-autofit-a-range-of-columns-to-observe-overridden-settings.cs
- create-a-custom-autofitteroptions-instance-that-disables-autofit-for-hidden-rows-and-apply-it-to-a-specific-row-range.cs
- after-autofitting-rows-export-the-worksheet-to-pdf-and-compare-file-size-with-a-nonfitted-version.cs
- set-row-height-for-the-header-row-to-improve-visual-emphasis-then-autofit-remaining-rows.cs
- autofit-all-rows-in-a-worksheet-by-calling-autofitrows-with-startrow-0-and-endrow-maxrow.cs
- apply-setcolumnwidthpixel-to-a-column-before-adding-numeric-data-to-control-column-width-precisely.cs
- load-an-xlsx-workbook-apply-autofitteroptionsforrendering-and-save-it-as-pdf.cs
- autofit-all-rows-for-normal-view-using-worksheetautofitrows-before-exporting-to-other-formats.cs
- autofit-rows-containing-merged-cells-to-ensure-merged-content-displays-correctly-in-pdf-output.cs
- enable-automatic-row-height-adjustment-on-workbook-load-to-preserve-original-layout.cs
- split-spacedelimited-text-in-column-a-into-separate-columns-using-texttocolumns.cs
- convert-commaseparated-values-in-column-b-to-individual-columns-by-specifying-comma-delimiter.cs
- parse-semicolonseparated-strings-in-a-column-using-texttocolumns-with-semicolon-delimiter.cs
- duplicate-multiple-consecutive-rows-using-copyrows-and-verify-formula-references-update-correctly.cs
- transfer-a-row-from-a-source-worksheet-to-a-destination-worksheet-using-cellscopyrow.cs
- copy-a-column-from-one-worksheet-to-another-while-maintaining-column-width-and-data-types.cs
- preserve-updated-formula-references-when-copying-rows-that-contain-relative-cell-references.cs
- disable-formula-adjustment-in-pasteoptions-to-copy-rows-with-absolute-references-unchanged.cs
- use-pasteoptions-to-copy-only-formatting-from-source-rows-excluding-values-and-formulas.cs
- apply-pasteoptions-to-copy-only-values-from-source-rows-ignoring-formulas-and-formatting.cs
- copy-rows-while-preserving-embedded-images-and-drawing-objects-using-default-copy-behavior.cs
- transfer-rows-containing-comments-and-verify-comments-appear-correctly-in-the-destination-worksheet.cs
- copy-rows-between-worksheets-then-autofit-destination-rows-to-match-source-row-heights.cs
- retrieve-source-row-height-with-getrowheight-and-explicitly-set-destination-height-using-setrowheight.cs
- batch-process-multiple-worksheets-by-autofitting-rows-for-rendering-before-saving-each-as-pdf.cs
- create-a-template-row-with-formulas-copy-it-to-several-sheets-and-validate-calculated-results.cs
- group-a-range-of-rows-programmatically-then-autofit-grouped-rows-to-ensure-proper-display.cs
- ungroup-previously-grouped-rows-and-verify-that-individual-row-heights-revert-to-original-values.cs
- group-columns-autofit-rows-within-grouped-columns-and-export-the-worksheet-to-pdf.cs
- ungroup-columns-after-modifications-and-ensure-column-widths-remain-consistent-across-the-worksheet.cs
- apply-autofitrows-after-inserting-new-data-rows-to-maintain-uniform-row-height-throughout-sheet.cs
- validate-that-autofitted-rows-do-not-exceed-a-maximum-height-limit-by-checking-row-heights.cs
- set-a-custom-maximum-row-height-before-autofitting-to-prevent-excessively-tall-rows.cs
- use-autofitrows-with-forrendering-on-a-worksheet-containing-wrapped-text-to-avoid-clipping.cs
- compare-row-heights-before-and-after-autofitrows-to-confirm-height-adjustments-were-applied.cs
- implement-error-handling-when-copying-rows-that-exceed-worksheet-row-limit-to-prevent-exceptions.cs
- copy-rows-while-preserving-hidden-row-states-ensuring-hidden-rows-remain-hidden-after-duplication.cs
- exclude-hidden-rows-from-copy-operation-by-filtering-rows-before-invoking-copyrows-method.cs
- copy-rows-that-are-part-of-an-autofiltered-range-and-verify-only-visible-rows-are-duplicated.cs
- after-copying-rows-recalculate-the-worksheet-to-update-any-dependent-formulas-automatically.cs
- use-worksheetcalculate-method-after-texttocolumns-operation-to-refresh-formulas-referencing-split-columns.cs
- validate-that-comments-attached-to-original-rows-are-correctly-transferred-to-copied-rows.cs
- preserve-cell-styles-when-copying-rows-by-enabling-style-preservation-in-pasteoptions.cs
- copy-rows-with-conditional-formatting-and-ensure-formatting-rules-apply-to-destination-rows.cs
- transfer-rows-containing-data-validation-lists-and-verify-validation-rules-remain-functional-after-copy.cs
- copy-rows-that-include-hyperlinks-and-confirm-hyperlinks-point-to-correct-targets-in-new-location.cs
- use-cellscopyrow-to-duplicate-a-header-row-and-then-freeze-the-copied-header-for-scrolling.cs
- after-copying-rows-apply-autofitrows-to-header-rows-separately-to-maintain-consistent-header-height.cs
- copy-rows-from-a-protected-worksheet-by-temporarily-disabling-protection-then-reenable-after-copy.cs
- programmatically-unprotect-a-worksheet-copy-rows-and-protect-worksheet-again-with-same-password.cs
- copy-rows-containing-merged-cells-and-verify-merged-regions-are-preserved-in-the-destination.cs
- after-copying-rows-adjust-column-widths-using-autofitcolumns-to-accommodate-newly-copied-data.cs
- batch-copy-a-template-row-to-multiple-target-rows-across-several-worksheets-using-a-loop.cs
- load-an-excel-workbook-hide-rows-five-through-ten-and-export-the-result-as-pdf.cs
- open-a-spreadsheet-hide-column-c-then-save-the-modified-file-in-pdf-format.cs
- read-a-workbook-unhide-row-twelve-with-height-twenty-points-and-generate-a-pdf.cs
- load-a-worksheet-unhide-column-b-specifying-width-fifty-points-then-export-to-pdf.cs
- open-an-excel-file-hide-rows-twenty-to-twentyfive-and-save-the-output-as-pdf.cs
- load-a-workbook-hide-columns-d-through-g-then-create-a-pdf-of-the-sheet.cs
- read-a-spreadsheet-unhide-rows-thirty-to-thirtyfive-with-default-height-and-export-pdf.cs
- open-a-workbook-unhide-columns-h-through-j-using-fiftypoint-width-then-save-pdf.cs
- load-an-excel-document-insert-a-new-row-at-index-fifteen-and-output-the-file-as-pdf.cs
- open-a-spreadsheet-insert-three-rows-starting-at-row-twenty-preserving-formatting-then-save-pdf.cs
- read-a-workbook-insert-five-rows-at-position-thirty-with-all-formatting-copied-and-generate-pdf.cs
- load-a-file-insert-a-new-column-at-index-three-then-export-the-worksheet-as-pdf.cs
- open-an-excel-workbook-delete-row-eight-and-save-the-updated-document-in-pdf-format.cs
- read-a-spreadsheet-delete-rows-ten-through-fifteen-then-create-a-pdf-of-the-cleaned-sheet.cs
- load-a-workbook-delete-column-f-and-export-the-resulting-file-as-pdf.cs
- open-a-worksheet-delete-all-blank-rows-then-save-the-compacted-workbook-as-pdf.cs
- read-an-excel-file-delete-all-blank-columns-and-generate-a-pdf-of-the-trimmed-sheet.cs
- load-a-workbook-remove-duplicate-rows-while-preserving-formulas-then-export-the-cleaned-file-as-pdf.cs
- open-a-spreadsheet-hide-rows-twentyone-to-twentyfive-then-save-the-view-as-pdf.cs
- read-a-workbook-hide-columns-k-through-m-and-produce-a-pdf-showing-hidden-columns.cs
- load-an-excel-file-unhide-rows-forty-to-fortyfive-with-custom-height-then-export-pdf.cs
- open-a-worksheet-unhide-columns-n-through-p-using-seventypoint-width-and-save-as-pdf.cs
- read-a-workbook-insert-two-rows-at-index-fifty-with-default-formatting-then-generate-pdf.cs
- load-a-spreadsheet-insert-a-column-at-position-five-and-copy-existing-formatting-then-export-pdf.cs
- open-an-excel-document-delete-rows-sixty-through-sixtyfive-and-save-the-result-as-pdf.cs
- read-a-workbook-delete-column-q-then-create-a-pdf-of-the-modified-worksheet.cs
- load-a-workbook-hide-rows-ten-to-twenty-then-unhide-rows-fifteen-to-eighteen-with-height-save-pdf.cs
- open-a-spreadsheet-hide-columns-five-through-eight-then-unhide-columns-six-and-seven-with-width-export-pdf.cs
- read-an-excel-workbook-insert-rows-at-index-twenty-with-formatting-copied-from-source-rows-then-generate-pdf.cs
- load-a-file-insert-a-column-at-position-two-copy-cell-styles-and-save-the-result-as-pdf.cs
- open-a-workbook-delete-rows-thirty-to-fortytwo-then-insert-five-new-rows-with-default-height-export-pdf.cs
- read-a-spreadsheet-delete-column-z-then-hide-rows-fifty-to-fiftyfive-and-produce-pdf.cs
- load-a-workbook-hide-rows-fifty-to-fiftyfive-then-export-the-filtered-view-as-pdf.cs
- open-a-workbook-unhide-all-rows-in-the-first-sheet-with-default-height-and-export-to-pdf.cs
- read-a-spreadsheet-unhide-all-columns-in-the-second-worksheet-using-default-width-then-generate-pdf.cs
- load-a-workbook-from-a-file-and-access-the-first-worksheet.cs
- access-the-second-worksheet-in-the-loaded-workbook-to-prepare-for-formula-verification.cs
- retrieve-the-formula-of-cell-e3-in-the-second-worksheet-before-deletion.cs
- create-a-deleteoptions-instance-to-control-reference-updating-behavior-during-deletion.cs
- set-the-deleteoptionsupdatereference-property-to-true-to-enable-formula-updates.cs
- verify-that-the-deleteoptionsupdatereference-property-is-set-to-true-before-deletion.cs
- invoke-deleteblankrowsandcolumns-on-the-first-worksheet-using-the-configured-deleteoptions.cs
- calculate-all-formulas-in-the-workbook-after-the-deletion-operation.cs
- read-the-formula-of-cell-e3-in-the-second-worksheet-after-deletion.cs
- verify-that-the-formula-now-references-the-first-worksheet-cell-a1.cs
- write-the-updated-formula-and-its-calculated-value-of-cell-e3-to-the-console.cs
- set-the-deleteoptionsupdatereference-property-to-false-to-preserve-original-formulas.cs
- verify-that-deleteoptionsupdatereference-property-is-false-before-the-second-deletion.cs
- invoke-deleteblankrowsandcolumns-on-the-first-worksheet-with-updatereference-disabled-to-preserve-formulas.cs
- calculate-all-formulas-in-the-workbook-after-the-second-deletion-operation.cs
- read-the-formula-of-cell-e3-in-the-second-worksheet-after-the-second-deletion.cs
- verify-that-the-formula-remains-unchanged-and-its-value-becomes-zero.cs
- write-the-unchanged-formula-and-zero-value-of-cell-e3-to-the-console.cs
- delete-blank-rows-on-the-first-worksheet-using-default-deleteoptions-without-updatereference.cs
- verify-that-formulas-in-other-worksheets-remain-unchanged-after-default-row-deletion.cs
- delete-blank-columns-on-the-first-worksheet-using-default-deleteoptions-without-updatereference.cs
- verify-that-formulas-in-other-worksheets-remain-unchanged-after-default-column-deletion.cs
- delete-blank-rows-on-the-first-worksheet-with-deleteoptionsupdatereference-set-to-true.cs
- verify-that-formulas-referencing-deleted-rows-are-updated-accordingly-after-row-deletion.cs
- delete-blank-columns-on-the-first-worksheet-with-deleteoptionsupdatereference-set-to-true.cs
- verify-that-formulas-referencing-deleted-columns-are-updated-accordingly-after-column-deletion.cs
