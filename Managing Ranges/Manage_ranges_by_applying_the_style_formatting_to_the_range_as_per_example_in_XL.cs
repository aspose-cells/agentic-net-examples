using System;
using Aspose.Cells;
using System.Drawing;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsRangeStyleDemo
{
    class Program
    {
        static void Main()
        {
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            Style style = workbook.CreateStyle();
            style.Pattern = BackgroundType.Solid;
            style.ForegroundColor = Color.LightGreen;
            style.Font.Color = Color.Red;
            style.Font.IsBold = true;
            style.Font.Size = 12;
            style.Font.Name = "Calibri";

            AsposeRange range = worksheet.Cells.CreateRange("A1", "C3");
            range.SetStyle(style);

            for (int row = 0; row < range.RowCount; row++)
            {
                for (int col = 0; col < range.ColumnCount; col++)
                {
                    range[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            workbook.Save("StyledRange.xlsx", SaveFormat.Xlsx);
        }
    }
}