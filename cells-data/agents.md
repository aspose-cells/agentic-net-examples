# Cells Data Examples

This folder contains **Aspose.Cells for .NET** code examples related to:

Cells Data


## Purpose

These examples demonstrate common **Aspose.Cells APIs** used when working with:

- Workbooks
- Worksheets
- Cells
- Formulas
- Charts
- Data operations


## Example Files

Each `.cs` file demonstrates a specific task related to **Cells Data**.

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
- load-an-excel-workbook-from-a-file-path-and-verify-it-opens-successfully.cs
- select-a-worksheet-by-name-and-obtain-its-cells-collection-for-further-operations.cs
- access-cell-b2-using-its-a1-style-name-set-a-numeric-value-and-save-the-workbook.cs
- access-a-cell-using-zerobased-row-and-column-indices-read-its-value-and-log-it.cs
- access-a-cell-by-its-numeric-index-within-the-cells-collection-modify-its-background-color-and-save.cs
- retrieve-the-worksheets-maximum-display-range-and-use-it-to-define-a-print-area.cs
- iterate-through-all-cells-in-the-maximum-display-range-to-count-nonempty-cells.cs
- convert-all-stringbased-numeric-values-in-the-entire-workbook-to-true-numbers-using-convertstringtonumericvalue.cs
- convert-string-numeric-values-only-in-the-first-worksheet-while-leaving-other-sheets-unchanged.cs
- convert-numeric-strings-within-range-a1c10-on-a-worksheet-and-verify-conversion.cs
- convert-numeric-strings-in-a-range-using-culturespecific-decimal-separators-and-ensure-correct-parsing.cs
- add-subtotals-summing-column-d-for-rows-2100-placing-results-at-each-groups-bottom.cs
- add-subtotals-that-count-entries-in-column-b-for-rows-5200-using-outline-grouping.cs
- add-subtotals-that-calculate-the-average-of-column-f-positioning-the-summary-rows-at-the-top.cs
- add-subtotals-using-the-max-function-on-column-c-with-summary-rows-placed-after-each-group.cs
- add-subtotals-using-the-stddev-function-on-column-h-and-set-summary-position-to-bottom.cs
- add-subtotals-using-the-var-function-on-column-i-with-outline-enabled-for-hierarchical-display.cs
- add-subtotals-using-the-product-function-on-column-j-placing-summary-rows-at-the-top-of-groups.cs
- add-subtotals-using-the-countnumbers-function-on-column-k-without-creating-outline-levels.cs
- add-subtotals-for-multiple-columns-simultaneously-summing-columns-m-and-n-together.cs
- apply-subtotals-after-sorting-the-worksheet-by-column-a-to-ensure-grouped-data-is-correct.cs
- apply-subtotals-on-filtered-data-confirming-that-only-visible-rows-are-included-in-calculations.cs
