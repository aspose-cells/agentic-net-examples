// Title: How to add a thick bottom border and light‑yellow fill to a cell with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set thick bottom border on a specific cell | C# Aspose.Cells apply solid light yellow background to a cell | Create custom cell style with border and fill using Aspose.Cells .NET | How to style cell A1 with border and background in Aspose.Cells C# | Set cell formatting (border style and fill color) in Aspose.Cells workbook
// Tags: Aspose.Cells cell style with border and fill | C# Aspose.Cells workbook cell formatting | set cell background color Aspose.Cells .NET | define custom cell style Aspose.Cells | apply border style to cell Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;

// The example creates a new workbook, defines a style that uses a thick bottom border and a solid light‑yellow fill, applies the style to cell A1, writes a value, and saves the file as StyledCell.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Create a new style
        Style style = workbook.CreateStyle();

        // Set a thick bottom border
        style.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thick;

        // Set light yellow fill color
        style.ForegroundColor = Color.LightYellow;
        style.Pattern = BackgroundType.Solid;

        // Apply the style to cell A1
        Cell cell = cells["A1"];
        cell.PutValue("Styled Cell");
        cell.SetStyle(style);

        // Save the workbook
        workbook.Save("StyledCell.xlsx");
    }
}
