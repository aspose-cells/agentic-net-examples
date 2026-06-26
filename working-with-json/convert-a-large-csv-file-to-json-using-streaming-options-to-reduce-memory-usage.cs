using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

class CsvToJsonStreaming
{
    static void Main()
    {
        try
        {
            // Paths for input CSV and output JSON
            string csvPath = "large_input.csv";
            string jsonOutputPath = "output.json";

            // Verify that the CSV file exists
            if (!File.Exists(csvPath))
            {
                Console.WriteLine($"Input file not found: {csvPath}");
                return;
            }

            // Create a new workbook and get the first worksheet's cells
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Open the CSV file as a stream to avoid loading the entire file into memory
            using (FileStream csvStream = new FileStream(csvPath, FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                // Import CSV data starting at cell A1 (row 0, column 0)
                // Use comma as delimiter and convert numeric data where possible
                cells.ImportCSV(csvStream, ",", true, 0, 0);
            }

            // Determine the used range after import
            int rows = cells.MaxDataRow + 1;      // MaxDataRow is zero‑based
            int cols = cells.MaxDataColumn + 1;   // MaxDataColumn is zero‑based
            Aspose.Cells.Range usedRange = cells.CreateRange(0, 0, rows, cols);

            // Configure JSON export options
            JsonSaveOptions jsonOptions = new JsonSaveOptions
            {
                ExportAsString = true,      // Export cell values as strings
                SkipEmptyRows = true,       // Omit empty rows
                HasHeaderRow = true,        // Treat first row as header
                ExportEmptyCells = false,   // Do not include empty cells
                ExportNestedStructure = false
            };

            // Export the used range to a JSON string
            string json = JsonUtility.ExportRangeToJson(usedRange, jsonOptions);

            // Write the JSON string to a file using a stream (low memory footprint)
            using (StreamWriter writer = new StreamWriter(jsonOutputPath, false, System.Text.Encoding.UTF8))
            {
                writer.Write(json);
            }

            Console.WriteLine("CSV has been successfully converted to JSON.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}