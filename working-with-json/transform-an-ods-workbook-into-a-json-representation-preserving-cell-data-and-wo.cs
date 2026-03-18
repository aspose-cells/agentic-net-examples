using System;
using Aspose.Cells;

class OdsToJsonConverter
{
    static void Main()
    {
        // Path to the source ODS file
        string odsFilePath = "input.ods";

        // Load the ODS workbook (Aspose.Cells detects format from extension)
        Workbook workbook = new Workbook(odsFilePath);

        // Configure JSON save options to preserve hierarchy and cell data
        JsonSaveOptions jsonOptions = new JsonSaveOptions
        {
            // Export the workbook as a JSON object even if there is only one worksheet
            AlwaysExportAsJsonObject = true,

            // Preserve worksheet hierarchy (parent‑child structure) in the JSON output
            ExportNestedStructure = true,

            // Include empty cells as null values
            ExportEmptyCells = true,

            // Do not skip empty rows so the original layout is kept
            SkipEmptyRows = false,

            // Convert to the Excel‑style JSON structure
            ToExcelStruct = true
        };

        // Path for the resulting JSON file
        string jsonOutputPath = "output.json";

        // Save the workbook as JSON using the configured options
        workbook.Save(jsonOutputPath, jsonOptions);
    }
}