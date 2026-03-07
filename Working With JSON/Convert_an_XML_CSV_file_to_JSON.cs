using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

class Program
{
    static void Main()
    {
        // Input CSV file path (replace with actual path)
        string csvPath = "input.csv";

        // Output JSON file path
        string jsonPath = "output.json";

        // Create a new workbook and get the first worksheet's cells
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Import CSV data into the worksheet starting at cell A1 (row 0, column 0)
        // Using comma as delimiter and converting numeric data where possible
        cells.ImportCSV(csvPath, ",", true, 0, 0);

        // Determine the used range of the imported data
        int lastRow = cells.MaxDataRow;          // zero‑based index of the last row with data
        int lastColumn = cells.MaxDataColumn;    // zero‑based index of the last column with data

        // Create a range that covers all imported cells
        Aspose.Cells.Range dataRange = cells.CreateRange(0, 0, lastRow + 1, lastColumn + 1);

        // Configure JSON export options
        JsonSaveOptions jsonOptions = new JsonSaveOptions
        {
            HasHeaderRow = true,          // treat the first row as header
            ExportEmptyCells = true       // include empty cells in the output
        };

        // Export the range to a JSON string
        string jsonOutput = JsonUtility.ExportRangeToJson(dataRange, jsonOptions);

        // Write the JSON string to the output file
        File.WriteAllText(jsonPath, jsonOutput);

        Console.WriteLine("CSV file has been successfully converted to JSON.");
    }
}