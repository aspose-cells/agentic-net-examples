// Title: How to print gridlines but hide row and column headings in an Aspose.Cells worksheet (C#)
// AI Prompts: Generate C# code that enables PrintGridlines and disables PrintHeadings on the active worksheet using Aspose.Cells. | Show the exact PageSetup property settings required to print only gridlines for a workbook saved as .xlsx. | Explain how to configure worksheet printing options to exclude row/column headings while keeping gridlines in Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# enable gridlines printing and turn off headings for the first worksheet | How to set PageSetup.PrintHeadings = false in Aspose.Cells .NET | Print only gridlines in Excel file using Aspose.Cells library | Disable row and column headings when exporting workbook with Aspose.Cells C# | Worksheet printing options Aspose.Cells hide headings keep gridlines
// Tags: gridlines printing Aspose.Cells C# | hide headings Aspose.Cells | worksheet PageSetup options Aspose.Cells | excel printing settings .NET Aspose.Cells | gridlines without headings Aspose.Cells

using Aspose.Cells;

// CREATE_WORKBOOK
var workbook = new Workbook();

// Get the first worksheet (current worksheet)
Worksheet sheet = workbook.Worksheets[0];

// Enable printing of gridlines
sheet.PageSetup.PrintGridlines = true;

// Disable printing of row and column headings
sheet.PageSetup.PrintHeadings = false;

// SAVE_WORKBOOK
workbook.Save("Result.xlsx");
