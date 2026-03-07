using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsExamples
{
    public class DbfCsvToJsonConverter
    {
        public static void Run(string csvFilePath, string jsonOutputPath)
        {
            // Create an empty workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Import the CSV file into the worksheet.
            // Using comma as delimiter, converting numeric data, starting at cell A1 (row 0, column 0).
            cells.ImportCSV(csvFilePath, ",", true, 0, 0);

            // Determine the used range (including header row).
            int totalRows = cells.MaxDataRow + 1;      // MaxDataRow is zero‑based
            int totalColumns = cells.MaxDataColumn + 1;
            Aspose.Cells.Range usedRange = cells.CreateRange(0, 0, totalRows, totalColumns);

            // Configure JSON export options.
            JsonSaveOptions jsonOptions = new JsonSaveOptions
            {
                ExportEmptyCells = true,   // include empty cells in the output
                HasHeaderRow = true,       // first row contains column names
                ExportNestedStructure = false
            };

            // Export the range to a JSON string.
            string jsonResult = JsonUtility.ExportRangeToJson(usedRange, jsonOptions);

            // Write the JSON string to the specified output file.
            File.WriteAllText(jsonOutputPath, jsonResult);

            Console.WriteLine($"CSV file \"{csvFilePath}\" has been converted to JSON and saved as \"{jsonOutputPath}\".");
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length < 2)
            {
                Console.WriteLine("Usage: AsposeCellsRunner <csvFilePath> <jsonOutputPath>");
                return;
            }

            string csvFilePath = args[0];
            string jsonOutputPath = args[1];

            DbfCsvToJsonConverter.Run(csvFilePath, jsonOutputPath);
        }
    }
}