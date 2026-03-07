using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace CsvToJsonApp
{
    class Program
    {
        static void Main(string[] args)
        {
            string csvPath = args.Length > 0 ? args[0] : "input.csv";
            string jsonPath = args.Length > 1 ? args[1] : "output.json";

            CsvToJsonConverter.ConvertCsvToJson(csvPath, jsonPath);
            Console.WriteLine($"Converted '{csvPath}' to '{jsonPath}'.");
        }
    }

    public static class CsvToJsonConverter
    {
        public static void ConvertCsvToJson(string csvPath, string jsonPath)
        {
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Csv);
            Workbook workbook = new Workbook(csvPath, loadOptions);
            Worksheet worksheet = workbook.Worksheets[0];
            Aspose.Cells.Range usedRange = worksheet.Cells.MaxDisplayRange;

            JsonSaveOptions jsonOptions = new JsonSaveOptions
            {
                HasHeaderRow = true,
                ExportEmptyCells = true
            };

            string jsonContent = JsonUtility.ExportRangeToJson(usedRange, jsonOptions);
            File.WriteAllText(jsonPath, jsonContent);
        }
    }
}