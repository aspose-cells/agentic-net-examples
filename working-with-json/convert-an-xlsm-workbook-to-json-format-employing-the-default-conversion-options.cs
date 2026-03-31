using System;
using Aspose.Cells;

class XlsmToJsonConverter
{
    static void Main()
    {
        // Path to the source XLSM workbook
        string sourcePath = "input.xlsm";

        // Desired output JSON file path
        string jsonPath = "output.json";

        // Load the XLSM workbook (default load options are used)
        Workbook workbook = new Workbook(sourcePath);

        // Create default JSON save options (no additional configuration)
        JsonSaveOptions jsonOptions = new JsonSaveOptions();

        // Save the workbook as JSON using the default options
        workbook.Save(jsonPath, jsonOptions);

        Console.WriteLine($"Workbook '{sourcePath}' has been converted to JSON at '{jsonPath}'.");
    }
}