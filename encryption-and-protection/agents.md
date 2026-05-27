# Encryption and Protection Examples

This folder contains **Aspose.Cells for .NET** code examples related to:

Encryption and Protection


## Purpose

These examples demonstrate common **Aspose.Cells APIs** used when working with:

- Workbooks
- Worksheets
- Cells
- Formulas
- Charts
- Data operations


## Example Files

Each `.cs` file demonstrates a specific task related to **Encryption and Protection**.

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
- detect-the-workbook-file-format-from-an-input-stream-and-log-the-identified-type.cs
- verify-whether-a-loaded-xls-file-is-encrypted-and-retrieve-its-encryption-algorithm.cs
- if-a-loaded-workbook-is-not-encrypted-apply-default-encryption-using-a-preset-password.cs
- detect-the-encryption-type-of-an-ods-file-and-display-the-algorithm-name.cs
- identify-weak-encryption-on-a-loaded-workbook-and-reencrypt-it-with-a-stronger-algorithm.cs
