using System;
using Aspose.Cells;
using System.Drawing;

class ApplyStyleToRange
{
    static void Main()
    {
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        cells["B2"].PutValue("Header");
        cells["C3"].PutValue(12345);
        cells["D4"].PutValue(DateTime.Now);

        var range = cells.CreateRange("B2", "D5");

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

        range.SetStyle(style);

        workbook.Save("StyledRange.xlsx", SaveFormat.Xlsx);
    }
}