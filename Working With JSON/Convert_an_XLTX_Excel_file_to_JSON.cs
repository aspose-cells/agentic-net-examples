using System;
using Aspose.Cells;

class ConvertXltxToJson
{
    static void Main()
    {
        // Path to the source XLTX file
        string sourcePath = "template.xltx";

        // Path where the JSON output will be saved
        string jsonPath = "output.json";

        // Load the XLTX workbook
        Workbook workbook = new Workbook(sourcePath);

        // Configure JSON save options
        JsonSaveOptions jsonOptions = new JsonSaveOptions
        {
            // Export the workbook as a JSON object even if there is only one worksheet
            AlwaysExportAsJsonObject = true,
            // Preserve the hierarchical structure of the workbook
            ExportNestedStructure = true,
            // Skip rows that contain no data
            SkipEmptyRows = true
        };

        // Save the workbook as JSON using the configured options
        workbook.Save(jsonPath, jsonOptions);

        Console.WriteLine($"Conversion completed successfully. JSON saved to '{jsonPath}'.");
    }
}