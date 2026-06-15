using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace CsvToJsonStreaming
{
    class Program
    {
        static void Main(string[] args)
        {
            // Paths for input CSV and output JSON
            string csvPath = "large_input.csv";
            string jsonPath = "output.json";

            // Ensure the CSV file exists
            if (!File.Exists(csvPath))
            {
                Console.WriteLine($"CSV file not found: {csvPath}");
                return;
            }

            try
            {
                // Open the CSV file as a read‑only stream to avoid loading the whole file into memory
                using (FileStream csvStream = new FileStream(csvPath, FileMode.Open, FileAccess.Read, FileShare.Read))
                {
                    // Create a new workbook and get the first worksheet's cells collection
                    Workbook workbook = new Workbook();
                    Worksheet worksheet = workbook.Worksheets[0];
                    Cells cells = worksheet.Cells;

                    // Import CSV data using the stream (delimiter is comma, convert numeric data)
                    cells.ImportCSV(csvStream, ",", true, 0, 0);

                    // Determine the used range of the imported data
                    int rows = cells.MaxDataRow + 1;      // MaxDataRow is zero‑based
                    int cols = cells.MaxDataColumn + 1;   // MaxDataColumn is zero‑based
                    Aspose.Cells.Range usedRange = cells.CreateRange(0, 0, rows, cols);

                    // Configure JSON export options
                    JsonSaveOptions jsonOptions = new JsonSaveOptions
                    {
                        HasHeaderRow = true,          // First row contains column names
                        ExportEmptyCells = true,      // Include empty cells as null
                        SkipEmptyRows = false,        // Keep empty rows in the output
                        ExportNestedStructure = false // Flat table structure
                    };

                    // Export the range to a JSON string
                    string jsonResult = JsonUtility.ExportRangeToJson(usedRange, jsonOptions);

                    // Write the JSON string to the output file
                    File.WriteAllText(jsonPath, jsonResult);
                    Console.WriteLine($"CSV successfully converted to JSON and saved at: {jsonPath}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}