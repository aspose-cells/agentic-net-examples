using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

class Program
{
    static void Main()
    {
        // Path to the source CSV file
        string csvPath = "input.csv";

        // Path where the resulting JSON will be saved
        string jsonPath = "output.json";

        // Load CSV data into a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;
        cells.ImportCSV(csvPath, ",", true, 0, 0);

        // Determine the used range
        int totalRows = cells.MaxDataRow + 1;
        int totalColumns = cells.MaxDataColumn + 1;
        Aspose.Cells.Range usedRange = cells.CreateRange(0, 0, totalRows, totalColumns);

        // Configure JSON export options
        JsonSaveOptions jsonOptions = new JsonSaveOptions
        {
            ExportEmptyCells = true,
            HasHeaderRow = true,
            ExportNestedStructure = false
        };

        // Export the range to a JSON string
        string jsonResult = JsonUtility.ExportRangeToJson(usedRange, jsonOptions);

        // Write the JSON string to the output file
        File.WriteAllText(jsonPath, jsonResult);

        Console.WriteLine($"CSV data has been converted to JSON and saved to '{jsonPath}'.");
    }
}