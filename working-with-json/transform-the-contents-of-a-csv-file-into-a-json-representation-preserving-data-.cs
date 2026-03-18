using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

class CsvToJsonConverter
{
    static void Main()
    {
        // Path to the source CSV file
        string csvPath = "input.csv";

        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Import CSV data into the worksheet.
        // Parameters: file name, delimiter, convert numeric data, start row, start column
        cells.ImportCSV(csvPath, ",", true, 0, 0);

        // Determine the range that contains the imported data
        Aspose.Cells.Range dataRange = cells.MaxDisplayRange;

        // Configure JSON export options
        JsonSaveOptions jsonOptions = new JsonSaveOptions
        {
            HasHeaderRow = true,          // First row contains column names
            ExportEmptyCells = true,      // Include empty cells in the output
            ExportNestedStructure = false // Flat table structure (CSV is not hierarchical)
        };

        // Export the range to a JSON string
        string jsonOutput = JsonUtility.ExportRangeToJson(dataRange, jsonOptions);

        // Write the JSON string to a file
        File.WriteAllText("output.json", jsonOutput);

        Console.WriteLine("CSV has been converted to JSON and saved as output.json");
    }
}