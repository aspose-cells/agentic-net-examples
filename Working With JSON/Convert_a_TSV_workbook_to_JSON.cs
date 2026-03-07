using System;
using Aspose.Cells;                     // Main Aspose.Cells namespace
using Aspose.Cells.Utility;            // For JsonUtility if needed (not used here)

class TsvToJsonConverter
{
    static void Main()
    {
        // Path to the source TSV file
        string tsvPath = "input.tsv";

        // Load the TSV workbook. Aspose.Cells automatically detects the format from the extension.
        Workbook workbook = new Workbook(tsvPath);

        // Configure JSON export options as needed.
        JsonSaveOptions jsonOptions = new JsonSaveOptions
        {
            // Export the workbook as a JSON object (default behavior)
            // You can customize additional options here, e.g.:
            // ExportNestedStructure = true,
            // SkipEmptyRows = true,
            // HasHeaderRow = true,
            // ExportAsString = true
        };

        // Path for the resulting JSON file
        string jsonPath = "output.json";

        // Save the workbook as JSON using the configured options.
        workbook.Save(jsonPath, jsonOptions);

        Console.WriteLine($"TSV file '{tsvPath}' has been converted to JSON at '{jsonPath}'.");
    }
}