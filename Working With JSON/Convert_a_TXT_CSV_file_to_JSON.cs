using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace CsvToJsonConversion
{
    class Program
    {
        static void Main()
        {
            string csvPath = "input.csv";

            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            worksheet.Cells.ImportCSV(csvPath, ",", true, 0, 0);

            Aspose.Cells.Range dataRange = worksheet.Cells.MaxDisplayRange;

            JsonSaveOptions jsonOptions = new JsonSaveOptions
            {
                ExportEmptyCells = true,
                HasHeaderRow = true,
                ExportNestedStructure = false
            };

            string jsonOutput = JsonUtility.ExportRangeToJson(dataRange, jsonOptions);

            string jsonPath = "output.json";
            File.WriteAllText(jsonPath, jsonOutput);

            Console.WriteLine($"CSV data from '{csvPath}' has been converted to JSON and saved to '{jsonPath}'.");
        }
    }
}