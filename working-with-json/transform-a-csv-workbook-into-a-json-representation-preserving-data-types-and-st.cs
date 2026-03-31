using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

class CsvToJsonConverter
{
    static void Main()
    {
        // Input CSV file path
        string csvPath = "input.csv";

        // Output JSON file path
        string jsonPath = "output.json";

        // Create a new workbook
        Workbook workbook = new Workbook();

        // Import CSV data into the first worksheet.
        // Using comma as delimiter, converting numeric strings to numeric types.
        workbook.Worksheets[0].Cells.ImportCSV(csvPath, ",", true, 0, 0);

        // Determine the used range (rows and columns with data)
        Worksheet sheet = workbook.Worksheets[0];
        int lastRow = sheet.Cells.MaxDataRow;      // zero‑based index of last data row
        int lastCol = sheet.Cells.MaxDataColumn;   // zero‑based index of last data column

        // Create a range that covers all imported data
        Aspose.Cells.Range dataRange = sheet.Cells.CreateRange(0, 0, lastRow + 1, lastCol + 1);

        // Configure JSON export options to preserve original data types
        JsonSaveOptions jsonOptions = new JsonSaveOptions
        {
            ExportAsString = false,      // keep numeric/date types, not forced to string
            ExportEmptyCells = false,    // omit empty cells
            HasHeaderRow = true,         // first row contains column names
            ExportNestedStructure = false,
            ToExcelStruct = false,
            Indent = "  "                // pretty‑print with two‑space indentation
        };

        // Export the range to a JSON string
        string json = JsonUtility.ExportRangeToJson(dataRange, jsonOptions);

        // Write the JSON string to a file
        System.IO.File.WriteAllText(jsonPath, json);

        Console.WriteLine("CSV has been successfully converted to JSON.");
    }
}