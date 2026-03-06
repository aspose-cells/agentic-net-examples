using System;
using System.Drawing;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

class Program
{
    static void Main()
    {
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        AsposeRange sourceRange = cells.CreateRange("A1:C3");
        for (int i = 0; i < sourceRange.RowCount; i++)
        {
            for (int j = 0; j < sourceRange.ColumnCount; j++)
            {
                sourceRange[i, j].PutValue($"R{i + 1}C{j + 1}");
            }
        }

        Style srcStyle = workbook.CreateStyle();
        srcStyle.Font.Name = "Arial";
        srcStyle.Font.Size = 12;
        srcStyle.Font.IsBold = true;
        srcStyle.ForegroundColor = Color.LightBlue;
        srcStyle.Pattern = BackgroundType.Solid;
        sourceRange.SetStyle(srcStyle);

        AsposeRange destRange = cells.CreateRange("E5:G7");
        sourceRange.Copy(destRange);

        workbook.Save("RangeCopyWithFormatting.xlsx");
    }
}