// Title: Print gridlines while suppressing row and column headings on the active worksheet using Aspose.Cells for .NET
// AI Prompts: Write C# code that sets PageSetup.PrintGridlines to true and PageSetup.PrintHeadings to false for the currently active worksheet in an Aspose.Cells workbook. | Show how to configure Aspose.Cells to print only gridlines and omit row/column headings when saving an Excel file. | Provide a step‑by‑step example that loads a workbook, modifies the active sheet’s print settings to enable gridlines and disable headings, and saves the result.
// Common Searches: Aspose.Cells .NET enable gridline printing and hide headings for active sheet | C# Aspose.Cells print only gridlines without row and column headings | How to set PrintHeadings false in Aspose.Cells page setup | Disable printing of row/column headers while keeping gridlines in Excel using Aspose.Cells | Aspose.Cells page setup print options for active worksheet
// Tags: Aspose.Cells PrintGridlines setting | Aspose.Cells PrintHeadings false | active worksheet page setup Aspose.Cells | disable Excel headings printing Aspose.Cells | enable gridline printing Aspose.Cells .NET

using Aspose.Cells;

// Load an existing workbook (replace with your file path)
Workbook workbook = new Workbook("input.xlsx");

// Get the currently active worksheet
Worksheet sheet = workbook.Worksheets[workbook.Worksheets.ActiveSheetIndex];

// Enable printing of gridlines
sheet.PageSetup.PrintGridlines = true;

// Disable printing of row and column headings
sheet.PageSetup.PrintHeadings = false;

// Save the modified workbook (replace with your desired output path)
workbook.Save("output.xlsx");
