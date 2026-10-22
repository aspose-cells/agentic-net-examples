// Title: Set a cell's horizontal alignment to Justify and enable text wrapping with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set cell justification and wrap text | apply justified horizontal alignment with text wrap in Excel using Aspose.Cells | C# code to enable text wrapping and justify alignment for a worksheet cell | how to format cell style for justify alignment and wrap in Aspose.Cells .NET | set cell style to Justify and wrap long text in an .xlsx file with Aspose.Cells
// Tags: Aspose.Cells cell justification C# | enable text wrapping Aspose.Cells | horizontal alignment Justify Aspose.Cells | cell style formatting .NET Excel | apply wrap and justify to Excel cell

using Aspose.Cells;
using System.Drawing;

// Create a new workbook
Workbook workbook = new Workbook();

// Access the first worksheet
Worksheet sheet = workbook.Worksheets[0];

// Get the cell you want to format (e.g., A1)
Cell cell = sheet.Cells["A1"];

// Create a new style or get the existing style of the cell
Style style = cell.GetStyle();

// Set horizontal alignment to Justify
style.HorizontalAlignment = TextAlignmentType.Justify;

// Enable text wrapping
style.IsTextWrapped = true;

// Apply the style back to the cell
cell.SetStyle(style);

// Optionally, put some long text to see the effect
cell.PutValue("This is a long piece of text that should be justified and wrapped within the cell to demonstrate the alignment and wrapping settings.");

// Save the workbook to a file
workbook.Save("AlignedWrappedCell.xlsx");
