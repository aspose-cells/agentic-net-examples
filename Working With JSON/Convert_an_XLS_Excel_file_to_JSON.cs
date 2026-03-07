using System;
using Aspose.Cells;

class ConvertXlsToJson
{
    static void Main()
    {
        // Source Excel file (XLS)
        string sourcePath = "input.xls";

        // Destination JSON file
        string jsonPath = "output.json";

        // Load the workbook from the XLS file
        Workbook workbook = new Workbook(sourcePath);

        // Configure JSON save options (optional customizations)
        JsonSaveOptions saveOptions = new JsonSaveOptions
        {
            // Export as a JSON object even if there is only one worksheet
            AlwaysExportAsJsonObject = true,
            // Include empty cells in the output
            ExportEmptyCells = true,
            // Treat the first row as header (if present)
            HasHeaderRow = true,
            // Export nested structure (worksheets as separate objects)
            ExportNestedStructure = true
        };

        // Save the workbook as JSON using the configured options
        workbook.Save(jsonPath, saveOptions);

        Console.WriteLine($"Conversion completed successfully. JSON saved to '{jsonPath}'.");
    }
}