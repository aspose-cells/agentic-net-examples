using System;
using System.IO;
using Aspose.Cells;

namespace NumbersToJsonConverter
{
    public class Converter
    {
        public static void ConvertWorkbookToJson(string workbookPath, string jsonOutputPath)
        {
            Workbook workbook;
            if (File.Exists(workbookPath))
            {
                workbook = new Workbook(workbookPath);
            }
            else
            {
                workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "Sheet1";
                sheet.Cells["A1"].PutValue(10);
                sheet.Cells["B1"].Formula = "=A1*2";
                sheet.Cells["C1"].PutValue("Text");
                workbook.Save(workbookPath);
            }

            workbook.CalculateFormula();

            JsonSaveOptions jsonOptions = new JsonSaveOptions();
            workbook.Save(jsonOutputPath, jsonOptions);
        }

        public static void Main()
        {
            string sourceWorkbook = Path.Combine(Environment.CurrentDirectory, "Sample.xlsx");
            string targetJson = Path.Combine(Environment.CurrentDirectory, "Sample.json");

            ConvertWorkbookToJson(sourceWorkbook, targetJson);

            Console.WriteLine("Conversion completed.");
        }
    }
}