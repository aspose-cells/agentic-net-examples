using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

class FodsToJson
{
    static void Main()
    {
        // Input FODS file path
        string inputPath = "input.fods";

        // Output JSON file path
        string outputPath = "output.json";

        // Load the FODS spreadsheet with the correct format
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Fods);
        Workbook workbook = new Workbook(inputPath, loadOptions);

        // Access the first worksheet (adjust if multiple sheets are needed)
        Worksheet sheet = workbook.Worksheets[0];

        // Determine the used range of the worksheet
        int firstRow = sheet.Cells.MinRow;
        int firstColumn = sheet.Cells.MinColumn;
        int totalRows = sheet.Cells.MaxRow - firstRow + 1;
        int totalColumns = sheet.Cells.MaxColumn - firstColumn + 1;

        // Create a range that covers all used cells
        Aspose.Cells.Range usedRange = sheet.Cells.CreateRange(firstRow, firstColumn, totalRows, totalColumns);

        // Configure JSON export options
        JsonSaveOptions jsonOptions = new JsonSaveOptions
        {
            ExportEmptyCells = true // keep empty cells in the output
            // ExportColumnHeaders is not supported in this version; headers will be included if present in the range
        };

        // Export the range to a JSON string
        string json = JsonUtility.ExportRangeToJson(usedRange, jsonOptions);

        // Write the JSON string to a file
        File.WriteAllText(outputPath, json);
    }
}