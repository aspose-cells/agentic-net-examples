// Title: How to apply an italic font and a thin top border to a cell using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells set cell font to italic and add top border | apply custom style with thin top border to a specific cell in Aspose.Cells | Aspose.Cells create style with italic text and top border in .NET | how to format cell A1 with italic font and thin top border using Aspose.Cells | set cell style borders and font properties with Aspose.Cells C#
// Tags: italic font styling Aspose.Cells | thin top border application Aspose.Cells | cell style definition C# Aspose.Cells | apply style to specific worksheet cell .NET | formatting cell borders using Aspose.Cells API

// Create a new workbook
var workbook = new Aspose.Cells.Workbook();

// Get the first worksheet
var worksheet = workbook.Worksheets[0];

// Create a new style object
var style = workbook.CreateStyle();

// Set the font to italic
style.Font.IsItalic = true;

// Apply a thin top border
style.Borders[Aspose.Cells.BorderType.TopBorder].LineStyle = Aspose.Cells.CellBorderType.Thin;

// Apply the style to a specific cell (e.g., A1)
var cell = worksheet.Cells["A1"];
cell.SetStyle(style);

// Save the workbook to a file
workbook.Save("StyledWorkbook.xlsx");
