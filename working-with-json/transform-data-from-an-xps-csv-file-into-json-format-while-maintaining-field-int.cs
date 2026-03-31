using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

class CsvToJsonConverter
{
    static void Main()
    {
        // Path to the source CSV file (XPS CSV)
        string csvPath = "input.csv";

        // Path where the resulting JSON will be saved
        string jsonPath = "output.json";

        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Import the CSV data starting at cell A1 (row 0, column 0)
        // Using comma as delimiter and converting numeric strings to numbers
        cells.ImportCSV(csvPath, ",", true, 0, 0);

        // Determine the used range of the imported data
        int lastRow = cells.MaxDataRow;          // zero‑based index of the last row with data
        int lastColumn = cells.MaxDataColumn;    // zero‑based index of the last column with data

        // Create a range that covers all imported cells
        Aspose.Cells.Range dataRange = cells.CreateRange(0, 0, lastRow + 1, lastColumn + 1);

        // Configure JSON export options
        JsonSaveOptions jsonOptions = new JsonSaveOptions
        {
            ExportEmptyCells = true,   // include empty cells in the output
            HasHeaderRow = true,       // treat the first row as header
            ExportNestedStructure = false // flat table structure
        };

        // Export the range to a JSON string
        string json = JsonUtility.ExportRangeToJson(dataRange, jsonOptions);

        // Write the JSON string to the output file
        File.WriteAllText(jsonPath, json);

        Console.WriteLine($"CSV data has been converted to JSON and saved to '{jsonPath}'.");
    }
}