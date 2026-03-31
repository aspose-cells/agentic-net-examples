using System;
using Aspose.Cells;

class XltxToJson
{
    static void Main()
    {
        // Path to the source XLTX template
        string sourcePath = "template.xltx";

        // Load the XLTX workbook
        Workbook workbook = new Workbook(sourcePath);

        // Configure JSON save options to preserve Excel structure
        JsonSaveOptions jsonOptions = new JsonSaveOptions
        {
            // Export as a JSON structure that mirrors the Excel workbook
            ToExcelStruct = true,
            // Ensure the output is a JSON object even if there is a single worksheet
            AlwaysExportAsJsonObject = true,
            // Include empty cells as null
            ExportEmptyCells = true,
            // Treat the first row as header (optional, adjust as needed)
            HasHeaderRow = true,
            // Export cell values as strings (optional)
            ExportAsString = true
        };

        // Destination JSON file path
        string jsonPath = "output.json";

        // Save the workbook as JSON using the configured options
        workbook.Save(jsonPath, jsonOptions);

        Console.WriteLine($"Workbook '{sourcePath}' has been converted to JSON at '{jsonPath}'.");
    }
}