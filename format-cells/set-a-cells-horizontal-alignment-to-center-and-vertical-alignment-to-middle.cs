// Title: How to center text horizontally and vertically in a single Excel cell using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that uses Aspose.Cells to set both the HorizontalAlignment and VerticalAlignment of cell A1 to Center and then saves the workbook. | Show how to retrieve a cell's Style, modify its TextAlignmentType properties for horizontal and vertical centering, and reapply the style with Aspose.Cells in C#.
// Common Searches: asp.net aspose.cells set cell A1 horizontal alignment to center | c# aspose.cells align cell content both horizontally and vertically | example code for centering text in a specific Excel cell using Aspose.Cells | how to apply vertical and horizontal text alignment with Aspose.Cells C#
// Tags: set cell horizontal alignment Aspose.Cells C# | set cell vertical alignment Aspose.Cells C# | Aspose.Cells cell style alignment example | center text in Excel cell using Aspose.Cells | Aspose.Cells TextAlignmentType usage

// Create a new workbook (lifecycle create rule)
Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook();

// Access the first worksheet
Aspose.Cells.Worksheet worksheet = workbook.Worksheets[0];

// Get the target cell (e.g., A1)
Aspose.Cells.Cell cell = worksheet.Cells["A1"];

// Retrieve the current style of the cell
Aspose.Cells.Style style = cell.GetStyle();

// Set horizontal alignment to center
style.HorizontalAlignment = Aspose.Cells.TextAlignmentType.Center;

// Set vertical alignment to middle (center)
style.VerticalAlignment = Aspose.Cells.TextAlignmentType.Center;

// Apply the modified style back to the cell
cell.SetStyle(style);

// Save the workbook (lifecycle save rule)
workbook.Save("AlignedCell.xlsx");
