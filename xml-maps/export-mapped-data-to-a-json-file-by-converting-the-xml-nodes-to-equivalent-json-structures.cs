using System;
using Aspose.Cells;

class ExportMappedDataToJson
{
    static void Main()
    {
        // Load the workbook that contains the XML map and the bound data
        Workbook workbook = new Workbook("MappedData.xlsx");

        // Ensure the workbook has at least one XML map
        if (workbook.Worksheets.XmlMaps.Count == 0)
        {
            Console.WriteLine("No XML map found in the workbook.");
            return;
        }

        // Retrieve the first XML map (adjust index if a specific map is required)
        XmlMap xmlMap = workbook.Worksheets.XmlMaps[0];

        // Configure JSON save options to produce a hierarchical (parent‑child) JSON structure
        JsonSaveOptions jsonOptions = new JsonSaveOptions
        {
            ExportNestedStructure = true,   // Export as nested JSON hierarchy
            HasHeaderRow = true,            // Treat the first row as header names
            ExportEmptyCells = true,        // Preserve empty cells as null in JSON
            AlwaysExportAsJsonObject = true,
            Indent = "  "                   // Pretty‑print with two‑space indentation
        };

        // Save the workbook as a JSON file; the data linked via the XML map will be represented in the JSON output
        string jsonOutputPath = "MappedData.json";
        workbook.Save(jsonOutputPath, jsonOptions);

        Console.WriteLine($"Mapped data successfully exported to JSON file: {jsonOutputPath}");
    }
}