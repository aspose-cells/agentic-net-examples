using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsPrnToJson
{
    class Program
    {
        static void Main()
        {
            string prnFilePath = "input.prn";

            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            cells.ImportCSV(prnFilePath, " ", true, 0, 0);

            Aspose.Cells.Range usedRange = worksheet.Cells.MaxDisplayRange;

            JsonSaveOptions jsonOptions = new JsonSaveOptions
            {
                ExportEmptyCells = true,
                HasHeaderRow = true
            };

            string jsonOutput = JsonUtility.ExportRangeToJson(usedRange, jsonOptions);

            string jsonFilePath = "output.json";
            File.WriteAllText(jsonFilePath, jsonOutput);

            Console.WriteLine($"PRN data has been converted to JSON and saved to '{jsonFilePath}'.");
        }
    }
}