using System;
using Aspose.Cells;

class ConvertXltmToJson
{
    static void Main()
    {
        // Path to the source XLTM workbook
        string sourcePath = "template.xltm";

        // Path for the resulting JSON file
        string jsonPath = "output.json";

        // Load the XLTM workbook (create & load lifecycle)
        Workbook workbook = new Workbook(sourcePath);

        // Configure JSON save options to preserve data types and sheet structure
        JsonSaveOptions jsonOptions = new JsonSaveOptions
        {
            // Export as a JSON object even if there is only one worksheet
            AlwaysExportAsJsonObject = true,

            // Preserve the Excel structural hierarchy (sheets, rows, columns)
            ToExcelStruct = true,

            // Include empty cells as null to keep the exact layout
            ExportEmptyCells = true,

            // Do not treat the first row as a header (preserve original rows)
            HasHeaderRow = false,

            // Export parent‑child hierarchy when applicable
            ExportNestedStructure = true,

            // Do not skip empty rows so row indices match the original sheet
            SkipEmptyRows = false
        };

        // Save the workbook as JSON using the configured options (save lifecycle)
        workbook.Save(jsonPath, jsonOptions);
    }
}