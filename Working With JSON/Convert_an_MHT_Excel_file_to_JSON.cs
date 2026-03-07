using System;
using Aspose.Cells;

class ConvertMhtToJson
{
    static void Main()
    {
        // Source MHT file path
        string sourcePath = "input.mht";

        // Destination JSON file path
        string outputPath = "output.json";

        if (!System.IO.File.Exists(sourcePath))
        {
            Console.WriteLine($"Source file not found: {sourcePath}");
            return;
        }

        // Load the MHT Excel file (auto detection)
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Auto);
        Workbook workbook = new Workbook(sourcePath, loadOptions);

        // Configure JSON save options (using defaults)
        JsonSaveOptions jsonOptions = new JsonSaveOptions();

        // Save the workbook as JSON
        workbook.Save(outputPath, jsonOptions);

        Console.WriteLine($"Conversion completed. JSON saved to: {outputPath}");
    }
}