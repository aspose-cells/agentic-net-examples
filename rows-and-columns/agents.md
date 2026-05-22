# Rows and Columns Examples

This folder contains **Aspose.Cells for .NET** code examples related to:

Rows and Columns


## Purpose

These examples demonstrate common **Aspose.Cells APIs** used when working with:

- Workbooks
- Worksheets
- Cells
- Formulas
- Charts
- Data operations


## Example Files

Each `.cs` file demonstrates a specific task related to **Rows and Columns**.

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
- load-a-workbook-with-loadoptionsautofitteroptionsonlyauto-set-to-true-to-automatically-adjust-all-row-heights.cs
- set-the-height-of-a-specific-row-eg-row-5-to-a-defined-point-value-using-cellssetrowheight.cs
- apply-a-uniform-row-height-to-all-rows-in-a-worksheet-by-assigning-cellsstandardheight.cs
- set-the-width-of-a-specific-column-eg-column-3-using-cellssetcolumnwidth.cs
- set-the-width-of-a-specific-column-in-pixels-using-cellssetcolumnwidthpixel.cs
- apply-a-uniform-column-width-to-all-columns-by-assigning-cellsstandardwidth.cs
- autofit-a-single-row-based-on-its-content-using-worksheetautofitrow.cs
- autofit-a-range-of-rows-eg-rows-1520-using-worksheetautofitrows.cs
- autofit-a-single-column-based-on-its-content-using-worksheetautofitcolumn.cs
