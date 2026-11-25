// Title: How to apply a thick bottom border and light‑yellow fill to a cell with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set thick bottom border for a specific cell | C# Aspose.Cells apply solid light yellow background to a cell | Create custom cell style with border and fill using Aspose.Cells .NET | Set style for cell A1 with bottom border and yellow fill in Aspose.Cells | Aspose.Cells example for styling cell borders and background color in C#
// Tags: Aspose.Cells cell border styling C# | Aspose.Cells solid fill color example | C# apply custom style to worksheet cell | Aspose.Cells thick bottom border setting | Aspose.Cells light yellow background formatting

// Create a new workbook
var workbook = new Aspose.Cells.Workbook();

// Access the first worksheet
var worksheet = workbook.Worksheets[0];

// Create a new style
var style = workbook.CreateStyle();

// Set a thick bottom border
style.Borders[Aspose.Cells.BorderType.BottomBorder].LineStyle = Aspose.Cells.CellBorderType.Thick;

// Set a light yellow fill color
style.ForegroundColor = System.Drawing.Color.LightYellow;
style.Pattern = Aspose.Cells.BackgroundType.Solid;

// Apply the style to a specific cell (e.g., A1)
var cell = worksheet.Cells["A1"];
cell.SetStyle(style);

// Save the workbook to a file
workbook.Save("StyledCell.xlsx");
