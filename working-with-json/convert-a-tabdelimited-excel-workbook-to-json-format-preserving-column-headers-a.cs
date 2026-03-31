using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

class TabToJsonConverter
{
    static void Main()
    {
        // Input TSV file and output JSON file paths
        string tsvPath = "input.tsv";
        string jsonPath = "output.json";

        // Load the TSV file using LoadOptions with Tsv format
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Tsv);
        Workbook workbook = new Workbook(tsvPath, loadOptions);

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Determine the used range (including headers)
        int firstRow = sheet.Cells.MaxDisplayRange.FirstRow;
        int firstColumn = sheet.Cells.MaxDisplayRange.FirstColumn;
        int totalRows = sheet.Cells.MaxDisplayRange.RowCount;
        int totalColumns = sheet.Cells.MaxDisplayRange.ColumnCount;

        // Create a range that covers the entire used area
        Aspose.Cells.Range range = sheet.Cells.CreateRange(firstRow, firstColumn, totalRows, totalColumns);

        // Configure JSON export options to preserve headers and data types
        JsonSaveOptions jsonOptions = new JsonSaveOptions
        {
            HasHeaderRow = true,
            ExportEmptyCells = true,
            ExportNestedStructure = false
        };

        // Export the range to a JSON string
        string json = JsonUtility.ExportRangeToJson(range, jsonOptions);

        // Write the JSON string to the output file
        File.WriteAllText(jsonPath, json);

        Console.WriteLine($"Conversion completed successfully. JSON saved to '{jsonPath}'.");
    }
}