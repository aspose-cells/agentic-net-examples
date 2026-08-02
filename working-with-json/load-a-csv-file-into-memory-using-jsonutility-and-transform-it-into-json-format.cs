using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

class Program
{
    static void Main()
    {
        try
        {
            // Path to the source CSV file
            string csvPath = "data.csv";

            // Verify that the CSV file exists to avoid FileNotFoundException
            if (!File.Exists(csvPath))
            {
                Console.WriteLine($"Error: CSV file not found at path '{csvPath}'.");
                return;
            }

            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Import CSV data into the worksheet starting at cell A1 (row 0, column 0)
            // Using comma as delimiter and converting numeric data where possible
            cells.ImportCSV(csvPath, ",", true, 0, 0);

            // Determine the used range dimensions
            int totalRows = cells.MaxRow + 1;      // MaxRow is zero‑based
            int totalColumns = cells.MaxColumn + 1;

            // Create a Range object that covers the used cells
            Aspose.Cells.Range usedRange = cells.CreateRange(0, 0, totalRows, totalColumns);

            // Configure JSON export options
            JsonSaveOptions jsonOptions = new JsonSaveOptions
            {
                ExportEmptyCells = true,
                HasHeaderRow = true,
                ExportNestedStructure = false
            };

            // Export the range to a JSON string
            string json = JsonUtility.ExportRangeToJson(usedRange, jsonOptions);

            // Output the JSON string to the console
            Console.WriteLine(json);

            // Optionally, write the JSON string to a file
            File.WriteAllText("output.json", json);
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}