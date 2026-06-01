using System;
using Aspose.Cells;

namespace WorksheetCalculateFormulaExternalDemo
{
    class Program
    {
        static void Main()
        {
            // ---------- Create external workbook with data ----------
            Workbook externalWb = new Workbook();
            // Set a file name that matches the reference used in formulas
            externalWb.FileName = "External.xlsx";

            // Populate data in Sheet1
            Worksheet extSheet1 = externalWb.Worksheets[0];
            extSheet1.Name = "Sheet1";
            extSheet1.Cells["A2"].PutValue(100);

            // Add a second sheet with additional data
            Worksheet extSheet2 = externalWb.Worksheets.Add("Sheet2");
            extSheet2.Cells["A2"].PutValue(200);

            // Define named ranges in the external workbook
            int nameIdx1 = externalWb.Worksheets.Names.Add("Sheet1!ExtNamedRange1");
            externalWb.Worksheets.Names[nameIdx1].RefersTo = "=Sheet1!$A$2";

            int nameIdx2 = externalWb.Worksheets.Names.Add("Sheet2!ExtNamedRange2");
            externalWb.Worksheets.Names[nameIdx2].RefersTo = "=Sheet2!$A$2";

            int nameIdx3 = externalWb.Worksheets.Names.Add("GlobalNamedRange");
            externalWb.Worksheets.Names[nameIdx3].RefersTo = "=Sheet1!$A$2";

            // ---------- Create main workbook that references the external one ----------
            Workbook mainWb = new Workbook();
            Worksheet mainSheet = mainWb.Worksheets[0];

            // Formulas that reference the external workbook
            mainSheet.Cells["A1"].Formula = "=[External.xlsx]Sheet1!$A$2";
            mainSheet.Cells["A2"].Formula = "=INDIRECT(\"[External.xlsx]Sheet1!$A$2\")";
            mainSheet.Cells["A3"].Formula = "=INDIRECT(\"[External.xlsx]Sheet1!ExtNamedRange1\")";
            mainSheet.Cells["A4"].Formula = "=INDIRECT(\"[External.xlsx]!GlobalNamedRange\")";

            // ---------- Set calculation options with linked data sources ----------
            CalculationOptions calcOptions = new CalculationOptions
            {
                // Provide the external workbook(s) that formulas may refer to
                LinkedDataSources = new Workbook[] { externalWb }
            };

            // Calculate all formulas in the worksheet, allowing recursive evaluation of external links
            mainSheet.CalculateFormula(calcOptions, true);

            // ---------- Output the calculated values ----------
            Console.WriteLine("Calculated values after Worksheet.CalculateFormula:");
            for (int i = 1; i <= 4; i++)
            {
                string cellName = $"A{i}";
                Console.WriteLine($"{cellName}: {mainSheet.Cells[cellName].StringValue}");
            }

            // (Optional) Save the main workbook to verify results in Excel
            // mainWb.Save("Result.xlsx", SaveFormat.Xlsx);
        }
    }
}