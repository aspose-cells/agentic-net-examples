using System;
using Aspose.Cells;               // Core Aspose.Cells namespace
using Aspose.Cells.Json;          // Namespace for JsonSaveOptions

class XlsToJsonConverter
{
    static void Main()
    {
        // Path to the source XLS workbook
        string sourcePath = "input.xls";

        // Desired output JSON file path
        string jsonPath = "output.json";

        // Load the XLS workbook
        Workbook workbook = new Workbook(sourcePath);

        // Configure JSON save options to preserve worksheet structure
        JsonSaveOptions jsonOptions = new JsonSaveOptions
        {
            // Export as a JSON representation of the Excel structure
            ToExcelStruct = true,

            // Ensure the output is always a JSON object even for a single sheet
            AlwaysExportAsJsonObject = true
        };

        // Save the workbook as JSON using the configured options
        workbook.Save(jsonPath, jsonOptions);

        Console.WriteLine($"Workbook '{sourcePath}' has been converted to JSON at '{jsonPath}'.");
    }
}