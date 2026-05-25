# Manage formulas Examples

This folder contains **Aspose.Cells for .NET** code examples related to:

Manage formulas


## Purpose

These examples demonstrate common **Aspose.Cells APIs** used when working with:

- Workbooks
- Worksheets
- Cells
- Formulas
- Charts
- Data operations


## Example Files

Each `.cs` file demonstrates a specific task related to **Manage formulas**.

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
- load-an-existing-workbook-then-assign-a-sum-formula-to-cell-b2-using-cellformula.cs
- verify-that-the-assigned-formula-uses-english-us-function-names-and-commas-as-argument-separators.cs
- loop-through-column-a-assign-incremental-row-references-in-a-multiplication-formula-and-store-results.cs
- create-a-shared-formula-for-range-c3c12-by-invoking-cellsetsharedformula-on-the-first-cell.cs
- configure-workbooksettingsmaxrowsofsharedformula-to-limit-shared-formula-rows-to-fifty-in-the-workbook.cs
- set-maxrowsofsharedformula-to-zero-to-disable-shared-formulas-for-the-entire-workbook.cs
- calculate-all-formulas-in-the-workbook-programmatically-using-workbookcalculate-and-retrieve-updated-cell-values.cs
- read-the-calculated-result-of-cell-d5-after-invoking-the-worksheetcalculate-method.cs
- use-calculateformula-method-to-evaluate-a-single-complex-formula-without-recalculating-the-whole-worksheet.cs
- retrieve-the-formula-string-from-cell-e10-and-log-it-for-audit-purposes.cs
- detect-cells-containing-calculation-errors-after-evaluation-and-log-their-addresses-for-investigation.cs
- replace-all-instances-of-the-today-function-with-a-static-date-string-to-freeze-calculations.cs
