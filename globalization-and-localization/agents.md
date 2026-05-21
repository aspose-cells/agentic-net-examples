# Globalization and Localization Examples

This folder contains **Aspose.Cells for .NET** code examples related to:

Globalization and Localization


## Purpose

These examples demonstrate common **Aspose.Cells APIs** used when working with:

- Workbooks
- Worksheets
- Cells
- Formulas
- Charts
- Data operations


## Example Files

Each `.cs` file demonstrates a specific task related to **Globalization and Localization**.

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
- create-a-custom-globalizationsettings-class-overriding-getlocalfunctionname-for-target-language-functions.cs
- override-geterrorstring-in-the-custom-class-to-provide-localized-error-messages-for-excel-errors.cs
- override-getbooleanstring-to-return-localized-truefalse-strings-for-the-selected-locale.cs
- assign-the-custom-globalizationsettings-instance-to-workbooksettingsglobalizationsettings-before-loading-any-worksheets.cs
- load-the-excel-workbook-using-workbookload-after-configuring-the-custom-globalization-settings.cs
- set-cell-formulas-with-cellformulalocal-to-apply-localized-function-names-during-workbook-processing.cs
- verify-that-localized-function-names-are-correctly-recognized-by-excel-when-the-workbook-is-opened.cs
- save-the-localized-workbook-as-xlsx-preserving-original-formatting-comments-and-cell-styles.cs
- generate-a-report-listing-processed-workbooks-applied-locales-and-any-localization-errors-encountered.cs
- implement-fallback-to-english-function-names-when-a-requested-locale-lacks-a-defined-mapping.cs
- log-each-overridden-method-call-to-a-debug-file-for-troubleshooting-localization-behavior-at-runtime.cs
- create-a-batch-process-that-applies-the-custom-globalization-settings-to-all-workbooks-in-a-folder.cs
- validate-that-boolean-values-display-localized-truefalse-strings-in-cells-containing-logical-formulas.cs
- write-unit-tests-asserting-getlocalfunctionname-returns-expected-localized-equivalents-for-common-functions.cs
- write-unit-tests-for-geterrorstring-covering-standard-excel-error-codes-across-multiple-locales.cs
- write-unit-tests-for-getbooleanstring-covering-true-false-and-null-values-in-different-locales.cs
- ensure-that-cell-comments-retain-their-original-language-while-function-names-are-localized-according-to-settings.cs
