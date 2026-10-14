// Title: Set the active worksheet to the third sheet (index 2) using Aspose.Cells for .NET in C#
// AI Prompts: Load an existing Excel workbook with Aspose.Cells, assign ActiveSheetIndex = 2, and save the updated file. | Programmatically make the third worksheet the active sheet in a workbook and write the result to a new file using the Aspose.Cells C# API.
// Common Searches: Aspose.Cells C# set active worksheet by zero based index | How to change the active sheet to the third tab in an Excel file using Aspose.Cells | C# Aspose.Cells ActiveSheetIndex example for selecting sheet 3 | Set workbook active sheet to index 2 with Aspose.Cells for .NET | Programmatically activate third worksheet in Excel using Aspose.Cells API
// Tags: Aspose.Cells set ActiveSheetIndex | C# select third worksheet Aspose.Cells | Excel workbook active sheet change .NET | zero based sheet index Aspose.Cells | activate specific worksheet C# Aspose.Cells

// Load an existing workbook
Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook("input.xlsx");

// Set the active worksheet to the third sheet (zero‑based index 2)
workbook.Worksheets.ActiveSheetIndex = 2;

// Save the workbook with the updated active sheet
workbook.Save("output.xlsx");
