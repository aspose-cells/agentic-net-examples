// Title: Hide an Excel table header row while keeping data rows visible using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells hide first row of an Excel table | How to hide Excel table header with Aspose.Cells without removing data rows | Aspose.Cells HideRow method to hide header row in a .xlsx file | Programmatically hide Excel header row while preserving row data using Aspose.Cells | Hide table header row in Excel workbook with Aspose.Cells C# code
// Tags: Aspose.Cells hide table header row | HideRow API C# Aspose.Cells | Excel worksheet row visibility Aspose.Cells | Preserve data rows while hiding header Aspose.Cells | Modify row visibility in .xlsx using Aspose.Cells

// Create a new workbook instance (or load an existing one)
Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook();

// Load the source workbook that contains the table
workbook = new Aspose.Cells.Workbook("input.xlsx");

// Access the first worksheet (adjust index if needed)
Aspose.Cells.Worksheet sheet = workbook.Worksheets[0];

// Hide the header row (row index 0). Data rows remain visible.
sheet.Cells.HideRow(0);

// Save the modified workbook
workbook.Save("output.xlsx");
