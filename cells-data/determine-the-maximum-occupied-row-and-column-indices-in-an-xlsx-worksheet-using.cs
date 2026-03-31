using System;
using Aspose.Cells;
using System.Drawing;
using AsposeRange = Aspose.Cells.Range;

class Program
{
    static void Main()
    {
        Workbook workbook = new Workbook("input.xlsx");
        Worksheet sheet = workbook.Worksheets[0];

        int maxRow = sheet.Cells.MaxDataRow;       
        int maxColumn = sheet.Cells.MaxDataColumn; 

        Console.WriteLine($"Maximum data row index (0‑based): {maxRow}");
        Console.WriteLine($"Maximum data column index (0‑based): {maxColumn}");

        if (maxRow >= 0 && maxColumn >= 0)
        {
            AsposeRange dataRange = sheet.Cells.CreateRange(0, 0, maxRow + 1, maxColumn + 1);

            Style style = workbook.CreateStyle();
            style.ForegroundColor = Color.LightYellow;
            style.Pattern = BackgroundType.Solid;

            StyleFlag flag = new StyleFlag { CellShading = true };
            dataRange.ApplyStyle(style, flag);
        }

        workbook.Save("output.xlsx");
    }
}