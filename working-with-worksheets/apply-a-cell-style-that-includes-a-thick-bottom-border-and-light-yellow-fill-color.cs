using System;
using System.Drawing;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Create a new style
        Style style = workbook.CreateStyle();

        // Set a light yellow fill
        style.Pattern = BackgroundType.Solid;
        style.ForegroundColor = Color.LightYellow;

        // Set a thick bottom border (black color)
        style.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thick;
        style.Borders[BorderType.BottomBorder].Color = Color.Black;

        // Apply the style to cell A1
        Cell cell = sheet.Cells["A1"];
        cell.PutValue("Styled Cell");
        cell.SetStyle(style);

        // Save the workbook
        workbook.Save("StyledCell.xlsx");
    }
}