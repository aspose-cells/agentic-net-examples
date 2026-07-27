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
- copy-a-worksheet-within-the-same-workbook-using-its-name-and-ensure-content-integrity.cs
- copy-multiple-worksheets-whose-names-start-with-a-specific-prefix-into-a-new-workbook-for-backup.cs
- addcopy-a-worksheet-and-specify-insertion-index-to-place-the-copy-directly-after-the-original-sheet.cs
- copy-a-worksheet-and-keep-all-conditional-formatting-rules-intact-for-consistent-styling.cs
- move-a-worksheet-to-a-new-position-by-providing-the-target-index-within-the-same-workbook.cs
- shift-a-worksheet-to-the-last-index-of-the-workbook-to-place-it-at-the-end.cs
- place-a-worksheet-immediately-after-a-specified-sheet-name-to-control-sheet-sequencing.cs
- enable-page-break-preview-mode-for-a-worksheet-to-visualize-printed-page-divisions.cs
- switch-a-worksheet-back-to-normal-view-mode-to-display-cells-without-page-break-outlines.cs
- apply-page-break-preview-to-every-worksheet-in-the-workbook-to-prepare-for-printing.cs
- set-all-worksheets-to-normal-view-mode-to-ensure-consistent-onscreen-display-across-the-workbook.cs
