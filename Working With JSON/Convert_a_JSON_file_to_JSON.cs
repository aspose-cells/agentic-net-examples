using System;
using Aspose.Cells;

class JsonToJsonConverter
{
    static void Main()
    {
        // Input and output JSON file paths
        string inputPath = "input.json";
        string outputPath = "output.json";

        // Create load options and keep the original schema
        JsonLoadOptions loadOptions = new JsonLoadOptions();
        loadOptions.KeptSchema = true;

        // Load the JSON file into a workbook using the load options
        Workbook workbook = new Workbook(inputPath, loadOptions);

        // Create save options for exporting back to JSON
        JsonSaveOptions saveOptions = new JsonSaveOptions();
        saveOptions.ExportNestedStructure = true; // Preserve hierarchical structure
        saveOptions.SkipEmptyRows = true;         // Omit empty rows in the output

        // Save the workbook as a JSON file using the save options
        workbook.Save(outputPath, saveOptions);

        Console.WriteLine("JSON file has been converted and saved successfully.");
    }
}