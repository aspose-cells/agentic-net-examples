# Manage Workbook Examples

This folder contains **Aspose.Cells for .NET** code examples related to:

Manage Workbook


## Purpose

These examples demonstrate common **Aspose.Cells APIs** used when working with:

- Workbooks
- Worksheets
- Cells
- Formulas
- Charts
- Data operations


## Example Files

Each `.cs` file demonstrates a specific task related to **Manage Workbook**.

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
- load-an-xlsx-workbook-from-a-file-stream-and-enable-automatic-formula-calculation.cs
- load-a-csv-file-into-a-workbook-specify-the-delimiter-and-treat-the-first-row-as-headers.cs
- load-a-workbook-from-a-memory-stream-modify-a-cell-value-and-write-back-to-the-stream.cs
- create-a-shared-workbook-instance-and-configure-it-for-concurrent-editing-by-multiple-users.cs
- set-workbook-calculation-engine-to-use-multithreaded-processing-for-faster-evaluation-of-large-data-sets.cs
- enable-iterative-calculation-mode-and-set-maximum-iterations-to-improve-convergence-of-circular-formulas.cs
