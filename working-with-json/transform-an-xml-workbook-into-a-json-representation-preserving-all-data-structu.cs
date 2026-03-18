using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Metadata;

class Program
{
    static void Main()
    {
        // Path to the source workbook (Excel XML format or any supported format)
        string sourcePath = "input.xml";          // replace with your workbook file
        // Path where the JSON representation will be saved
        string jsonPath = "output.json";

        // Load the workbook
        Workbook workbook = new Workbook(sourcePath);

        // Configure JSON save options to preserve the Excel structure and metadata
        JsonSaveOptions jsonOptions = new JsonSaveOptions
        {
            // Export as a JSON object even if there is only one worksheet
            AlwaysExportAsJsonObject = true,
            // Preserve parent‑child hierarchy (nested structure)
            ExportNestedStructure = true,
            // Convert the Excel structure (tables, merged cells, etc.) into JSON
            ToExcelStruct = true,
            // Keep empty cells as null so the layout is unchanged
            ExportEmptyCells = true,
            // Do not skip empty rows – they are part of the original layout
            SkipEmptyRows = false,
            // Header row handling – set according to your data (false keeps raw rows)
            HasHeaderRow = false
        };

        // Save the workbook as JSON
        workbook.Save(jsonPath, jsonOptions);

        // ----- Append workbook metadata (document properties) to the JSON file -----
        // Load metadata (built‑in and custom document properties)
        MetadataOptions metaOptions = new MetadataOptions(MetadataType.DocumentProperties);
        WorkbookMetadata metadata = new WorkbookMetadata(sourcePath, metaOptions);

        // Build a simple JSON fragment for metadata
        using (StreamWriter writer = new StreamWriter(jsonPath, true))
        {
            writer.WriteLine(","); // separate from the previous JSON content
            writer.WriteLine("\"Metadata\": {");

            // Built‑in document properties
            writer.WriteLine("  \"BuiltIn\": {");
            bool firstProp = true;
            foreach (var prop in metadata.BuiltInDocumentProperties)
            {
                if (!firstProp) writer.WriteLine(",");
                writer.Write($"    \"{prop.Name}\": \"{prop.Value}\"");
                firstProp = false;
            }
            writer.WriteLine();
            writer.WriteLine("  },");

            // Custom document properties
            writer.WriteLine("  \"Custom\": {");
            firstProp = true;
            foreach (var prop in metadata.CustomDocumentProperties)
            {
                if (!firstProp) writer.WriteLine(",");
                writer.Write($"    \"{prop.Name}\": \"{prop.Value}\"");
                firstProp = false;
            }
            writer.WriteLine();
            writer.WriteLine("  }");

            writer.WriteLine("}"); // end of Metadata object
        }

        // Clean up
        workbook.Dispose();

        Console.WriteLine("Workbook successfully converted to JSON with metadata.");
    }
}