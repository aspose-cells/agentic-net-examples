using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsExamples
{
    public class NumbersCsvToJsonConverter
    {
        public static void Run()
        {
            string csvPath = Path.Combine(Environment.CurrentDirectory, "numbers.csv");

            if (!File.Exists(csvPath))
            {
                File.WriteAllText(csvPath,
                    "Name,Age,Score\n" +
                    "Alice,30,85.5\n" +
                    "Bob,25,92.0\n" +
                    "Charlie,28,78.3");
            }

            string jsonOutputPath = Path.Combine(Environment.CurrentDirectory, "output.json");

            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            cells.ImportCSV(csvPath, ",", true, 0, 0);

            int totalRows = cells.MaxDataRow + 1;
            int totalColumns = cells.MaxDataColumn + 1;

            AsposeRange dataRange = cells.CreateRange(0, 0, totalRows, totalColumns);

            JsonSaveOptions jsonOptions = new JsonSaveOptions
            {
                ExportAsString = true,
                HasHeaderRow = true,
                ExportEmptyCells = true
            };

            string jsonResult = JsonUtility.ExportRangeToJson(dataRange, jsonOptions);

            File.WriteAllText(jsonOutputPath, jsonResult);

            Console.WriteLine($"CSV data from '{csvPath}' has been converted to JSON and saved to '{jsonOutputPath}'.");
        }

        public static void Main(string[] args)
        {
            Run();
        }
    }
}