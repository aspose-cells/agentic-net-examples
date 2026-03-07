using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsExamples
{
    public class CsvToJsonConverter
    {
        public static void Run()
        {
            // Paths for the source CSV, temporary XLSX, and final JSON output
            string csvPath = "input.csv";
            string tempXlsxPath = "temp.xlsx";
            string jsonOutputPath = "output.json";

            // Ensure a sample CSV exists (replace with your actual file as needed)
            if (!File.Exists(csvPath))
            {
                File.WriteAllText(csvPath, "Name,Age,City\nJohn,30,New York\nAlice,25,London");
            }

            // 1. Convert CSV to XLSX using the provided ConversionUtility method
            ConversionUtility.Convert(csvPath, tempXlsxPath);

            // 2. Load the converted workbook
            Workbook workbook = new Workbook(tempXlsxPath);

            // 3. Define the range to export – here we use the used range of the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            int firstRow = sheet.Cells.MinRow;
            int firstColumn = sheet.Cells.MinColumn;
            int totalRows = sheet.Cells.MaxRow - firstRow + 1;
            int totalColumns = sheet.Cells.MaxColumn - firstColumn + 1;
            Aspose.Cells.Range exportRange = sheet.Cells.CreateRange(firstRow, firstColumn, totalRows, totalColumns);

            // 4. Configure JSON save options (optional settings)
            JsonSaveOptions jsonOptions = new JsonSaveOptions
            {
                AlwaysExportAsJsonObject = true,
                HasHeaderRow = true,
                ExportEmptyCells = true
            };

            // 5. Export the range to a JSON string using the provided JsonUtility method
            string jsonResult = JsonUtility.ExportRangeToJson(exportRange, jsonOptions);

            // 6. Write the JSON string to the output file
            File.WriteAllText(jsonOutputPath, jsonResult);

            Console.WriteLine($"CSV file '{csvPath}' has been converted to JSON and saved as '{jsonOutputPath}'.");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            CsvToJsonConverter.Run();
        }
    }
}