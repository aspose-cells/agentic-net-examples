using System;
using Aspose.Cells;
using System.Drawing;
using AsposeRange = Aspose.Cells.Range;

class StyleRangeDemo
{
    static void Main()
    {
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        Style style = workbook.CreateStyle();
        style.Font.Name = "Calibri";
        style.Font.Size = 12;
        style.Font.IsBold = true;
        style.Font.Color = Color.White;
        style.ForegroundColor = Color.DarkBlue;
        style.Pattern = BackgroundType.Solid;
        style.HorizontalAlignment = TextAlignmentType.Center;
        style.VerticalAlignment = TextAlignmentType.Center;
        style.IsTextWrapped = true;
        style.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thick;
        style.Borders[BorderType.BottomBorder].Color = Color.Yellow;

        AsposeRange range = sheet.Cells.CreateRange("A1:C3");
        range.SetStyle(style);

        for (int row = 0; row < 3; row++)
        {
            for (int col = 0; col < 3; col++)
            {
                sheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
            }
        }

        workbook.Save("StyledRange.xlsx");
    }
}