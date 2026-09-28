// Title: Use a named range as the print area in Aspose.Cells for .NET and export only that range to PDF
// AI Prompts: Assign an existing named range to the worksheet's PageSetup.PrintArea and save the workbook as a PDF with Aspose.Cells in C#. | If the named range is missing, add it to the worksheet, set it as the print area, and generate a PDF that includes just that range using Aspose.Cells.
// Common Searches: Aspose.Cells C# set print area to a named range before PDF export | How to export only a named range to PDF with Aspose.Cells .NET | Limit PDF output to specific cells using named range in Aspose.Cells | C# example for setting worksheet print area to a named range and saving as PDF
// Tags: Aspose.Cells set print area named range | Aspose.Cells export named range to PDF | C# worksheet print area Aspose.Cells | PDF conversion limited to range Aspose.Cells | named range creation Aspose.Cells .NET

// Load the workbook from an existing Excel file
Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook("input.xlsx");

// Define the name of the range that should be used as the print area
string namedRange = "MyPrintRange";

// Retrieve the first worksheet (or specify the appropriate index/name)
Aspose.Cells.Worksheet sheet = workbook.Worksheets[0];

// Set the print area of the worksheet to the named range
// The PrintArea property expects the range in A1 style; using the named range name works directly
sheet.PageSetup.PrintArea = namedRange;

// Optionally, ensure the named range exists; if not, you can create it like this:
// workbook.Worksheets.Names.Add(namedRange, sheet.Name + "!A1:C10");

// Convert the workbook to PDF, which will now include only the cells within the specified print area
workbook.Save("output.pdf", Aspose.Cells.SaveFormat.Pdf);
