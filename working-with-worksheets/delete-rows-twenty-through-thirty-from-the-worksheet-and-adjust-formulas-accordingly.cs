// Title: Delete rows 20‑30 from the first worksheet and automatically adjust formulas with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that removes rows 20 through 30 from the first sheet of an XLSX workbook using Aspose.Cells and ensures all dependent formulas are recalculated. | Generate a snippet that loads an Excel file, deletes a specific range of rows with zero‑based indexing, and saves the file while preserving formula integrity using Aspose.Cells.
// Common Searches: Aspose.Cells C# delete rows 20-30 and keep formulas intact | how to remove a block of rows in Excel with Aspose.Cells and update formulas | C# Aspose.Cells delete multiple rows without breaking cell references
// Tags: delete rows Aspose.Cells C# | adjust formulas after row deletion Aspose.Cells | zero‑based row indexing Aspose.Cells | remove multiple rows Excel .NET | preserve formula references Aspose.Cells

using Aspose.Cells;

// Load the workbook
Workbook workbook = new Workbook("input.xlsx");

// Access the first worksheet
Worksheet ws = workbook.Worksheets[0];

// Delete rows 20 through 30 (inclusive).
// Aspose.Cells uses zero‑based indexing, so start at row 19 and delete 11 rows.
ws.Cells.DeleteRows(19, 11);

// Save the modified workbook
workbook.Save("output.xlsx");
