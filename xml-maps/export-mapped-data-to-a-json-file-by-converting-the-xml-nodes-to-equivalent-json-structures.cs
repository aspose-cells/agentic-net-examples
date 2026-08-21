// Title: Export XML‑Mapped Excel Data to JSON with Aspose.Cells for .NET (C#)
// Description: Loads an Excel workbook that contains an XML map, validates the map's presence, configures JsonSaveOptions (ExportNestedStructure, AlwaysExportAsJsonObject, HasHeaderRow, ExportEmptyCells, Indent) to preserve the XML hierarchy and empty cells, and saves the mapped data as a readable JSON file.
// Keywords: Aspose.Cells | C# | XML map | JSON export | JsonSaveOptions | ExportNestedStructure | AlwaysExportAsJsonObject | HasHeaderRow | ExportEmptyCells | Indent | Excel to JSON conversion | mapped worksheet data
// Common Searches: Aspose.Cells export XML map to JSON C# | How to save mapped Excel data as JSON using Aspose | JsonSaveOptions nested structure example | Convert Excel XML map to hierarchical JSON | C# export empty Excel cells as null in JSON
// Developer Intent: Convert an Excel workbook with an XML map into a structured JSON file using Aspose.Cells for .NET.
// Use Cases: Generate JSON payloads for APIs from Excel sheets already aligned with an XML schema. | Create configuration or settings files by turning mapped rows into nested JSON objects. | Produce human‑readable JSON reports that retain empty cells as null values.
// AI Prompts: Write C# code that loads an Excel file containing an XML map and exports the mapped data to JSON with hierarchical structure and nulls for empty cells using Aspose.Cells. | Explain the effect of each JsonSaveOptions property when exporting XML‑mapped data to JSON. | Add robust error handling for scenarios where the workbook lacks an XML map before attempting JSON conversion.

using System;
using Aspose.Cells;
using Aspose.Cells.Json; // JsonSaveOptions resides in this namespace

// Loads an Excel workbook that contains an XML map, validates the map's presence, configures JsonSaveOptions (ExportNestedStructure, AlwaysExportAsJsonObject, HasHeaderRow, ExportEmptyCells, Indent) to preserve the XML hierarchy and empty cells, and saves the mapped data as a readable JSON file.
class ExportMappedDataToJson
{
    static void Main()
    {
        // Load the workbook that contains the XML map and the mapped data
        Workbook workbook = new Workbook("MappedData.xlsx"); // replace with your file path

        // Ensure the workbook actually has an XML map; otherwise there is nothing to export
        if (workbook.Worksheets.XmlMaps.Count == 0)
        {
            Console.WriteLine("No XmlMap found in the workbook.");
            return;
        }

        // Configure JSON export options
        JsonSaveOptions jsonOptions = new JsonSaveOptions
        {
            // Export as a parent‑child hierarchy to reflect the XML structure
            ExportNestedStructure = true,

            // Always output a JSON object even if there is only one worksheet
            AlwaysExportAsJsonObject = true,

            // Treat the first row as header names (optional, based on your data)
            HasHeaderRow = true,

            // Include empty cells as null values
            ExportEmptyCells = true,

            // Indent the output for readability
            Indent = "  "
        };

        // Save the workbook as a JSON file using the configured options
        string outputPath = "MappedData.json";
        workbook.Save(outputPath, jsonOptions);

        Console.WriteLine($"Mapped data successfully exported to JSON file: {outputPath}");
    }
}
