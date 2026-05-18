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
- define-a-semicolon-as-csv-delimiter-in-jsonlayoutoptions-before-converting-json-to-csv.cs
- load-a-csv-file-into-memory-using-jsonutility-and-transform-it-into-json-format.cs
- export-a-specific-cell-range-from-a-worksheet-to-json-using-exportrangetojsonoptions.cs
- include-column-names-as-keys-in-json-output-by-enabling-includecolumnnames-option.cs
- set-json-output-encoding-to-utf-8-within-jsonsaveoptions-before-saving-workbook-as-json.cs
- batch-process-a-folder-of-xls-files-converting-each-workbook-to-separate-json-files.cs
- validate-json-structure-against-a-predefined-schema-after-loading-with-jsonutility-successfully.cs
- merge-multiple-json-files-into-a-single-workbook-creating-separate-worksheets-for-each-file.cs
- preserve-cell-formulas-during-json-conversion-by-configuring-jsonsaveoptionspreserveformulas-flag.cs
- generate-prettyprinted-json-with-indentation-by-setting-jsonsaveoptionsprettyprint-to-true.cs
- load-a-json-array-representing-multiple-tables-and-map-each-element-to-a-separate-worksheet.cs
- encrypt-the-generated-json-file-using-a-passwordprotected-stream-before-writing-to-disk.cs
- implement-error-handling-to-catch-jsonutilityload-exceptions-when-source-json-file-is-malformed.cs
- create-a-console-application-that-prints-json-representation-of-the-first-worksheet-to-standard-output.cs
- configure-jsonsaveoptions-to-exclude-hidden-rows-and-columns-from-the-exported-json-data.cs
- transform-numeric-values-to-strings-in-json-output-by-applying-a-custom-value-formatter.cs
