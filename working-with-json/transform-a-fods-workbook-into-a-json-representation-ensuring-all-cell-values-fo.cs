using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

class FodsToJsonConverter
{
    static void Main()
    {
        // Path to the source FODS file
        string fodsPath = "input.fods";

        // Load the FODS workbook with appropriate load options
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Fods);
        Workbook workbook = new Workbook(fodsPath, loadOptions);

        // Access the first worksheet (you can iterate if needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Determine the used range of the worksheet
        int maxRow = worksheet.Cells.MaxDataRow;      // zero‑based index of last row with data
        int maxCol = worksheet.Cells.MaxDataColumn;   // zero‑based index of last column with data

        // Create a range that covers all used cells
        Aspose.Cells.Range usedRange = worksheet.Cells.CreateRange(0, 0, maxRow + 1, maxCol + 1);

        // Configure JSON export options to include structure, empty cells, and header row
        JsonSaveOptions jsonOptions = new JsonSaveOptions
        {
            ToExcelStruct = true,          // export as Excel‑style JSON structure
            ExportEmptyCells = true,       // include empty cells in the output
            HasHeaderRow = true            // treat the first row as header (optional)
        };

        // Export the range to a JSON string
        string jsonOutput = JsonUtility.ExportRangeToJson(usedRange, jsonOptions);

        // Write the JSON string to a file
        string jsonPath = "output.json";
        File.WriteAllText(jsonPath, jsonOutput);

        // Optionally display the JSON on the console
        Console.WriteLine("JSON export completed. Output written to " + jsonPath);
        Console.WriteLine(jsonOutput);
    }
}