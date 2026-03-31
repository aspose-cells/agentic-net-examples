using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

class XlsToJsonConverter
{
    static void Main()
    {
        // Path to the source XLS file
        string sourcePath = "input.xls";

        // Path where the JSON output will be saved
        string jsonPath = "output.json";

        // Load the workbook from the XLS file
        Workbook workbook = new Workbook(sourcePath);

        // Configure JSON save options to preserve the full Excel structure,
        // cell types, empty cells, and ensure the result is a JSON object.
        JsonSaveOptions jsonOptions = new JsonSaveOptions();
        jsonOptions.AlwaysExportAsJsonObject = true;   // Export as JSON object even if only one sheet
        jsonOptions.ExportNestedStructure = true;      // Keep worksheet hierarchy in JSON
        jsonOptions.ExportEmptyCells = true;           // Include empty cells in the output
        jsonOptions.ToExcelStruct = true;              // Convert to Excel struct format

        // Save the workbook as a JSON file using the configured options
        workbook.Save(jsonPath, jsonOptions);

        Console.WriteLine($"Conversion completed successfully. JSON saved to '{jsonPath}'.");
    }
}