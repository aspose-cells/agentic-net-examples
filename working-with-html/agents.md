# Working With HTML Examples

This folder contains **Aspose.Cells for .NET** code examples related to:

Working With HTML


## Purpose

These examples demonstrate common **Aspose.Cells APIs** used when working with:

- Workbooks
- Worksheets
- Cells
- Formulas
- Charts
- Data operations


## Example Files

Each `.cs` file demonstrates a specific task related to **Working With HTML**.

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
- set-htmlsaveoptions-to-disable-scientific-notation-for-numbers-exceeding-a-specified-threshold.cs
- configure-workbooks-numberformat-to-display-full-integer-values-when-exporting-to-html.cs
- apply-custom-number-format-strings-to-ensure-large-identifiers-appear-without-scientific-notation.cs
- verify-that-html-output-shows-large-numbers-as-plain-text-rather-than-exponential-representation.cs
- use-the-saveoptions-property-to-enforce-nonexponential-display-for-all-numeric-cells.cs
- enable-the-option-to-keep-original-numeric-formatting-when-converting-excel-to-html.cs
- test-html-export-with-values-larger-than-10⁶-to-confirm-exponential-notation-is-suppressed.cs
- compare-exported-html-with-and-without-exponential-notation-suppression-for-accuracy.cs
