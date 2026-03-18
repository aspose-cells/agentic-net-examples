using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Utility;
using AsposeRange = Aspose.Cells.Range;

namespace SpreadsheetToJsonDemo
{
    class Program
    {
        static void Main()
        {
            string inputPath = "input.xlsx";
            Workbook workbook = new Workbook(inputPath);

            List<string> worksheetJsonList = new List<string>();

            foreach (Worksheet sheet in workbook.Worksheets)
            {
                int maxRow = sheet.Cells.MaxDataRow;
                int maxColumn = sheet.Cells.MaxDataColumn;

                if (maxRow < 0 || maxColumn < 0)
                {
                    worksheetJsonList.Add($"{{\"Name\":\"{sheet.Name}\",\"Data\":null}}");
                    continue;
                }

                AsposeRange usedRange = sheet.Cells.CreateRange(0, 0, maxRow + 1, maxColumn + 1);

                JsonSaveOptions jsonOptions = new JsonSaveOptions
                {
                    ToExcelStruct = true,
                    AlwaysExportAsJsonObject = true
                };

                string sheetJson = JsonUtility.ExportRangeToJson(usedRange, jsonOptions);

                worksheetJsonList.Add($"{{\"Name\":\"{sheet.Name}\",\"Data\":{sheetJson}}}");
            }

            string workbookJson = $"{{\"Worksheets\":[{string.Join(",", worksheetJsonList)}]}}";

            Console.WriteLine(workbookJson);
        }
    }
}