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
