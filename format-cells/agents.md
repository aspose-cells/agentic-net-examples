# Format Cells Examples

This folder contains **Aspose.Cells for .NET** code examples related to:

Format Cells


## Purpose

These examples demonstrate common **Aspose.Cells APIs** used when working with:

- Workbooks
- Worksheets
- Cells
- Formulas
- Charts
- Data operations


## Example Files

Each `.cs` file demonstrates a specific task related to **Format Cells**.

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
- load-an-xlsx-workbook-verify-theme-usage-and-extract-all-theme-color-definitions.cs
- create-a-custom-color-scheme-assigning-new-rgb-values-to-accent1-and-accent2-theme-colors.cs
- apply-the-custom-accent1-theme-color-to-the-background-of-cells-in-the-first-worksheet.cs
- set-the-custom-accent2-theme-color-as-the-font-color-for-header-rows-across-all-worksheets.cs
- retrieve-the-current-rgb-value-of-the-hyperlink-theme-color-from-a-loaded-workbook.cs
- update-the-hyperlink-theme-color-to-a-new-shade-of-blue-and-save-the-workbook.cs
- validate-that-cells-using-accent1-automatically-reflect-the-updated-theme-color-after-modification.cs
- clone-the-theme-from-a-template-workbook-and-assign-it-to-a-newly-created-workbook.cs
- check-whether-a-workbook-contains-theme-colors-before-applying-bulk-theme-updates-to-multiple-files.cs
- iterate-through-a-folder-of-xlsx-files-replace-accent3-with-a-custom-green-shade-and-save.cs
- generate-a-report-listing-each-workbooks-theme-color-palette-before-and-after-applying-changes.cs
- use-the-api-to-enumerate-all-theme-color-types-and-output-their-default-rgb-values.cs
- apply-a-conditional-formatting-rule-that-colors-cells-based-on-the-accent4-theme-color-intensity.cs
- create-a-chart-and-set-its-series-colors-to-use-the-workbooks-theme-accent5-colors.cs
- assign-the-themes-dark1-color-to-the-fill-of-a-pivot-tables-row-headers.cs
- programmatically-reset-a-workbooks-theme-to-the-default-office-theme-and-verify-cell-colors.cs
- load-a-workbook-disable-theme-usage-and-convert-all-themed-cells-to-explicit-rgb-formatting.cs
- measure-the-time-taken-to-update-all-theme-colors-in-a-large-workbook-with-thousands-of-cells.cs
- apply-a-custom-theme-color-to-the-border-of-a-range-of-cells-in-the-second-worksheet.cs
- extract-the-themes-font-scheme-and-list-the-primary-and-secondary-font-families-used.cs
- replace-the-workbooks-theme-font-with-a-custom-font-family-and-update-all-cell-styles.cs
- check-if-any-cells-use-the-themecolorindex-enumeration-and-log-their-addresses-for-review.cs
- create-a-new-workbook-assign-a-custom-theme-and-populate-it-with-sample-data-using-theme-colors.cs
- export-the-theme-color-palette-to-a-json-file-for-external-analysis-and-documentation.cs
- import-a-json-defined-color-scheme-and-apply-it-as-the-workbooks-new-theme-colors.cs
- detect-if-a-workbook-uses-theme-colors-and-fallback-to-direct-rgb-formatting-when-absent.cs
- update-the-themes-hyperlink-color-to-a-dark-gray-and-ensure-all-links-display-correctly.cs
- programmatically-compare-two-workbooks-theme-palettes-and-highlight-differences-in-a-report.cs
- load-a-workbook-apply-a-custom-theme-and-generate-a-pdf-preview-of-the-first-sheet.cs
- iterate-over-all-worksheets-set-each-header-rows-font-color-to-the-themes-accent1.cs
- create-a-macrofree-template-workbook-with-predefined-theme-colors-for-downstream-processing.cs
- validate-that-after-changing-the-theme-conditional-formatting-rules-still-reference-correct-color-indices.cs
- apply-a-gradient-fill-to-a-range-using-the-themes-accent3-and-accent4-colors.cs
- set-the-themes-dark2-color-as-the-default-border-color-for-all-tables-in-the-workbook.cs
- export-the-modified-workbook-to-xlsx-and-verify-that-the-theme-xml-reflects-new-colors.cs
