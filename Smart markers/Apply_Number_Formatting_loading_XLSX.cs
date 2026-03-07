using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        LoadOptions loadOptions = new LoadOptions();
        Workbook workbook = new Workbook("input.xlsx", loadOptions);
        Worksheet sheet = workbook.Worksheets[0];

        sheet.Cells["A1"].PutValue(1234.567);
        sheet.Cells["A2"].PutValue(0.5);

        Style numberStyle = workbook.CreateStyle();
        numberStyle.Custom = "0.00";

        StyleFlag flag = new StyleFlag();
        flag.NumberFormat = true;

        Aspose.Cells.Range targetRange = sheet.Cells.CreateRange("A1:A2");
        targetRange.ApplyStyle(numberStyle, flag);

        workbook.Save("output.xlsx");
    }
}