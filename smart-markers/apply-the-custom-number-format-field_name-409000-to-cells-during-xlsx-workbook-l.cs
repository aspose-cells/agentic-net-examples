using System;
using System.Globalization;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        string sourcePath = "input.xlsx";

        LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
        loadOptions.CultureInfo = new CultureInfo("en-US");

        Workbook workbook = new Workbook(sourcePath, loadOptions);

        Style customStyle = workbook.CreateStyle();
        customStyle.Custom = "[$-409]#,##0.00";

        StyleFlag flag = new StyleFlag();
        flag.NumberFormat = true;

        Worksheet sheet = workbook.Worksheets[0];
        int maxRow = sheet.Cells.MaxDataRow;
        int maxCol = sheet.Cells.MaxDataColumn;

        Aspose.Cells.Range usedRange = sheet.Cells.CreateRange(0, 0, maxRow + 1, maxCol + 1);
        usedRange.ApplyStyle(customStyle, flag);

        string outputPath = "output.xlsx";
        workbook.Save(outputPath, SaveFormat.Xlsx);
    }
}