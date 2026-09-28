// Title: How to add a thick left border and light‑blue background fill to a single cell with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that creates a style with a solid light‑blue fill and a thick left border, then applies it to cell A1 in a new Aspose.Cells workbook. | Modify the example to apply the same style to an entire range (e.g., B2:D5) while preserving the left‑border thickness and background color. | Extend the style to include thick borders on all four sides, keep the light‑blue fill, and apply it to cell C3.
// Common Searches: Aspose.Cells C# set left border thickness and background color for a cell | C# example of applying solid fill and custom border to a specific cell using Aspose.Cells | How to style cell A1 with light blue background and thick left border in Aspose.Cells .NET | Apply custom cell style to a range of cells in Aspose.Cells C# tutorial
// Tags: Aspose.Cells style left border thickness | light blue cell fill Aspose.Cells | C# apply custom style to cell A1 | Aspose.Cells solid background fill example | cell formatting borders fill Aspose.Cells .NET

// Create a new workbook
var workbook = new Aspose.Cells.Workbook();

// Access the first worksheet
var worksheet = workbook.Worksheets[0];

// Create a new style object
var style = workbook.CreateStyle();

// Set the background fill to light blue
style.ForegroundColor = System.Drawing.Color.LightBlue;
style.Pattern = Aspose.Cells.BackgroundType.Solid;

// Set a thick left border
style.Borders[Aspose.Cells.BorderType.LeftBorder].LineStyle = Aspose.Cells.CellBorderType.Thick;

// Apply the style to a specific cell (e.g., A1)
var cell = worksheet.Cells["A1"];
cell.SetStyle(style);

// Save the workbook to a file
workbook.Save("StyledWorkbook.xlsx");
