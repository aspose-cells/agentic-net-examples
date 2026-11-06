// Title: How to set a cell's horizontal alignment to justify and enable text wrapping using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set cell justification and wrap text | C# code to apply justify alignment and enable text wrapping for an Excel cell with Aspose.Cells | How to enable text wrap and justify alignment on a specific cell using Aspose.Cells .NET | Set horizontal alignment to Justify and wrap text in Aspose.Cells workbook C# | Apply justify alignment and text wrap style to cell A1 with Aspose.Cells
// Tags: justify horizontal alignment Aspose.Cells C# | enable text wrapping cell style Aspose.Cells | set cell alignment and wrap text Aspose.Cells | apply justify alignment to Excel cell using Aspose.Cells | format cell justification and wrap .NET

using Aspose.Cells;

// Assume workbook is created or loaded according to the lifecycle rules
Workbook workbook = new Workbook();               // create a new workbook
Worksheet sheet = workbook.Worksheets[0];        // get the first worksheet

// Target cell (e.g., A1)
Cell cell = sheet.Cells["A1"];

// Retrieve the current style of the cell
Style style = cell.GetStyle();

// Set horizontal alignment to Justify
style.HorizontalAlignment = TextAlignmentType.Justify;

// Enable text wrapping
style.IsTextWrapped = true;

// Apply the modified style back to the cell
cell.SetStyle(style);

// Save the workbook according to the lifecycle rules
workbook.Save("output.xlsx");   // save the workbook to a file
