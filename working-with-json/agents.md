# Working With JSON Examples

This folder contains **Aspose.Cells for .NET** code examples related to:

Working With JSON


## Purpose

These examples demonstrate common **Aspose.Cells APIs** used when working with:

- Workbooks
- Worksheets
- Cells
- Formulas
- Charts
- Data operations


## Example Files

Each `.cs` file demonstrates a specific task related to **Working With JSON**.

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
- load-an-xlsx-workbook-from-a-file-path-and-verify-successful-initialization.cs
- export-the-active-worksheet-of-the-loaded-workbook-to-json-using-default-saveformat-settings.cs
- export-the-entire-workbook-to-json-with-column-headers-included-via-jsonsaveoptions.cs
- specify-a-custom-date-format-in-jsonsaveoptions-before-exporting-workbook-to-json.cs
- exclude-empty-rows-from-json-output-by-setting-jsonsaveoptionsincludeemptyrows-to-false.cs
- convert-a-json-file-containing-tabular-data-to-csv-using-jsonutility-with-a-custom-delimiter.cs
