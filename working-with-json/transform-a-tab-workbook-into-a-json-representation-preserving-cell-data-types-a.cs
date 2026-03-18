using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsJsonExport
{
    class Program
    {
        static void Main(string[] args)
        {
            // Load an existing workbook (the TAB workbook) from file.
            // The constructor with a file path handles the loading lifecycle.
            Workbook workbook = new Workbook("input.xlsx");

            // Configure JSON save options to preserve cell data types,
            // keep empty cells, and maintain worksheet hierarchy.
            JsonSaveOptions jsonOptions = new JsonSaveOptions
            {
                // Export as a JSON object even if there is only one worksheet.
                AlwaysExportAsJsonObject = true,
                // Preserve the parent‑child (worksheet) structure.
                ExportNestedStructure = true,
                // Include empty cells in the output (null values).
                ExportEmptyCells = true,
                // Do not force all values to strings; keep original types.
                ExportAsString = false,
                // Optional: indent the JSON for readability.
                Indent = "  "
            };

            // Save the workbook as a JSON file using the configured options.
            // The Save method with JsonSaveOptions follows the required save lifecycle.
            workbook.Save("output.json", jsonOptions);
        }
    }
}