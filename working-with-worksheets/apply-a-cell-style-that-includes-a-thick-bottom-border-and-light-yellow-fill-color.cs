// Title: Aspose.Cells .NET: Add Thick Bottom Border & Light Yellow Fill to a Cell
// Description: Shows how to create a workbook, define a style with a solid light‑yellow background and a thick bottom border, apply the style to cell A1, and save the file as StyledCell.xlsx using C#.
// Keywords: Aspose.Cells C# style example | Excel cell border thickness | light yellow cell fill | solid background Aspose.Cells | GitHub Aspose.Cells sample | cell formatting .NET
// Common Searches: Aspose.Cells set thick bottom border | C# change cell fill color to yellow | how to style a single cell in Aspose.Cells | apply custom border and background in Excel with Aspose | example code for cell style Aspose.Cells .NET
// Developer Intent: Create a reusable style that adds a thick bottom edge and a light‑yellow background to a cell, then apply it in a workbook.
// Use Cases: Highlight header rows in automated reports with a bold bottom line and yellow shading. | Emphasize key totals in financial sheets by giving them a distinct border and fill. | Design a title cell for exported Excel files that stands out visually.
// AI Prompts: Generate C# code using Aspose.Cells to apply a thick bottom border and light yellow fill to the range B2:D2. | Show how to define a reusable Style with a thick bottom border and yellow background and apply it to multiple cells in an Aspose.Cells workbook. | Explain how to change the border color while keeping the thick bottom border and solid fill using Aspose.Cells for .NET.

using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsStyleExample
{
    // Shows how to create a workbook, define a style with a solid light‑yellow background and a thick bottom border, apply the style to cell A1, and save the file as StyledCell.xlsx using C#.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle create)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Create a style object
            Style style = workbook.CreateStyle();

            // Set a solid fill with light yellow background
            style.Pattern = BackgroundType.Solid;
            style.ForegroundColor = Color.LightYellow;

            // Configure the bottom border to be thick and black (or any color you prefer)
            style.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thick;
            style.Borders[BorderType.BottomBorder].Color = Color.Black;

            // Apply the style to a specific cell (e.g., A1)
            Cell cell = cells["A1"];
            cell.PutValue("Styled Cell");
            cell.SetStyle(style);

            // Save the workbook (lifecycle save)
            workbook.Save("StyledCell.xlsx");
        }
    }
}
