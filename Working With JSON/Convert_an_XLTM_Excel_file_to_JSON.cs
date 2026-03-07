using System;
using Aspose.Cells;

class ConvertXltmToJson
{
    static void Main()
    {
        // Path to the source XLTM file
        string sourcePath = "template.xltm";

        // Path where the JSON output will be saved
        string jsonPath = "output.json";

        // Load the XLTM workbook
        Workbook workbook = new Workbook(sourcePath);

        // Configure JSON save options
        JsonSaveOptions jsonOptions = new JsonSaveOptions();
        // Export as a JSON object even if the workbook contains a single worksheet
        jsonOptions.AlwaysExportAsJsonObject = true;

        // Save the workbook as JSON
        workbook.Save(jsonPath, jsonOptions);

        Console.WriteLine($"Conversion completed. JSON saved to: {jsonPath}");
    }
}