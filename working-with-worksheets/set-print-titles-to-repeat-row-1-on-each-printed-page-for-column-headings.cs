// Title: How to repeat the first worksheet row as a print title on every printed page using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that sets PageSetup.PrintTitleRows to "$1:$1" so the first row repeats on each printed page in an Aspose.Cells workbook. | Provide an example that also configures PageSetup.PrintTitleColumns to repeat the first column together with the header row. | Show how to save the workbook after applying print title settings with Aspose.Cells in C#.
// Common Searches: Aspose.Cells C# repeat header row on each printed page | set print title rows in Aspose.Cells workbook .NET | how to make first row print on every page using Aspose.Cells | Aspose.Cells page setup repeat first column and row when printing | C# Aspose.Cells print title rows syntax
// Tags: Aspose.Cells PageSetup.PrintTitleRows | repeat header row Aspose.Cells | print title rows Excel .NET | Aspose.Cells repeat first column print | C# workbook save Aspose.Cells example

// Create a new workbook (or load an existing one)
// Assuming a create rule exists for workbook initialization
Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook();

// Access the first worksheet (or any target worksheet)
Aspose.Cells.Worksheet sheet = workbook.Worksheets[0];

// Set the print title to repeat the first row on every printed page
// This makes row 1 act as a column heading on each page
sheet.PageSetup.PrintTitleRows = "$1:$1";

// (Optional) If you also want to repeat the first column, you could set:
// sheet.PageSetup.PrintTitleColumns = "$A:$A";

// Save the workbook (using the provided save rule)
workbook.Save("PrintTitlesExample.xlsx");
