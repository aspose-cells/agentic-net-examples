using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace NumbersToJsonExample
{
    class Program
    {
        static void Main()
        {
            string sourcePath = "input.numbers";

            Workbook workbook;
            if (File.Exists(sourcePath))
            {
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Numbers);
                workbook = new Workbook(sourcePath, loadOptions);
            }
            else
            {
                workbook = new Workbook();
                Worksheet ws = workbook.Worksheets[0];
                ws.Name = "Sheet1";
                ws.Cells["A1"].PutValue("Sample");
                ws.Cells["B1"].PutValue(123);
                ws.Cells["A2"].PutValue(DateTime.Now);
            }

            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                Worksheet sheet = workbook.Worksheets[i];

                int firstRow = sheet.Cells.MinRow;
                int firstColumn = sheet.Cells.MinColumn;
                int totalRows = sheet.Cells.MaxRow - firstRow + 1;
                int totalColumns = sheet.Cells.MaxColumn - firstColumn + 1;

                var usedRange = sheet.Cells.CreateRange(firstRow, firstColumn, totalRows, totalColumns);

                JsonSaveOptions jsonOptions = new JsonSaveOptions
                {
                    ToExcelStruct = true,
                    ExportEmptyCells = true,
                    HasHeaderRow = false
                };

                string json = JsonUtility.ExportRangeToJson(usedRange, jsonOptions);

                string outputPath = $"Sheet{i + 1}_{sheet.Name}.json";
                File.WriteAllText(outputPath, json);
                Console.WriteLine($"Worksheet '{sheet.Name}' exported to '{outputPath}'.");
            }

            workbook.Dispose();
        }
    }
}