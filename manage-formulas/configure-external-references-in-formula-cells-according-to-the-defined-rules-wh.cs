using System;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        string mainPath = Path.Combine(Directory.GetCurrentDirectory(), "Main.xlsx");
        string externalPath = Path.Combine(Directory.GetCurrentDirectory(), "External.xlsx");

        if (!File.Exists(mainPath))
        {
            var wb = new Workbook();
            var ws = wb.Worksheets[0];
            ws.Name = "Sheet1";
            ws.Cells["A1"].Formula = "=[External.xlsx]Sheet1!A1";
            wb.Save(mainPath);
        }

        if (!File.Exists(externalPath))
        {
            var wbExt = new Workbook();
            var wsExt = wbExt.Worksheets[0];
            wsExt.Name = "Sheet1";
            wsExt.Cells["A1"].PutValue(123);
            wbExt.Save(externalPath);
        }

        var mainWorkbook = new Workbook(mainPath);
        var externalWorkbook = new Workbook(externalPath);

        var options = new CalculationOptions();
        options.LinkedDataSources = new Workbook[] { externalWorkbook };

        mainWorkbook.CalculateFormula(options);

        string resultPath = Path.Combine(Directory.GetCurrentDirectory(), "Result.xlsx");
        mainWorkbook.Save(resultPath);
    }
}