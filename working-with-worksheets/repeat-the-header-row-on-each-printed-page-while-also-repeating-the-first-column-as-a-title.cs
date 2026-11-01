// Title: How to repeat the header row and first column on every printed page using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that sets PrintTitleRows and PrintTitleColumns so the first row and first column repeat on each printed page with Aspose.Cells. | Describe the steps to configure the page setup in Aspose.Cells to make both the header row and leftmost column appear on every printed sheet. | Provide a complete example that loads an existing Excel file, applies print title settings for rows and columns, and saves the workbook using Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# repeat first row on each printed page | Set print title column A in Aspose.Cells workbook | Configure page setup print titles for rows and columns in Aspose.Cells .NET | C# code to make header row and first column repeat when printing Excel with Aspose.Cells
// Tags: Aspose.Cells set PrintTitleRows C# | Aspose.Cells repeat header row print titles | Aspose.Cells PrintTitleColumns configuration | Aspose.Cells page setup repeat first column | C# Aspose.Cells print titles for Excel worksheets

// Load an existing workbook
Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook("input.xlsx");

// Access the first worksheet (or any specific worksheet)
Aspose.Cells.Worksheet sheet = workbook.Worksheets[0];

// Repeat the header row (first row) on every printed page
sheet.PageSetup.PrintTitleRows = "$1:$1";

// Repeat the first column as a title on every printed page
sheet.PageSetup.PrintTitleColumns = "$A:$A";

// Save the modified workbook
workbook.Save("output.xlsx");
