# Working With Worksheets Examples

This folder contains **Aspose.Cells for .NET** code examples related to:

Working With Worksheets


## Purpose

These examples demonstrate common **Aspose.Cells APIs** used when working with:

- Workbooks
- Worksheets
- Cells
- Formulas
- Charts
- Data operations


## Example Files

Each `.cs` file demonstrates a specific task related to **Working With Worksheets**.

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
- copy-a-worksheet-within-the-same-workbook-using-its-numeric-index-and-verify-duplication.cs
- copy-a-worksheet-within-the-same-workbook-using-its-name-and-ensure-content-integrity.cs
- copy-a-worksheet-from-a-source-workbook-to-a-target-workbook-while-preserving-formulas.cs
- copy-multiple-worksheets-whose-names-start-with-a-specific-prefix-into-a-new-workbook-for-backup.cs
- use-worksheetcollectionaddcopy-overload-with-source-index-to-duplicate-a-sheet-at-a-specific-position.cs
- use-worksheetcollectionaddcopy-overload-by-name-to-copy-a-worksheet-and-insert-it-after-a-target-sheet.cs
- addcopy-a-worksheet-and-specify-insertion-index-to-place-the-copy-directly-after-the-original-sheet.cs
