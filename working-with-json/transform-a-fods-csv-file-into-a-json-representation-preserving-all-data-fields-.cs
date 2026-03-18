using System;
using Aspose.Cells;               // Core Aspose.Cells namespace
using Aspose.Cells.Utility;      // For utility classes (if needed)

class Program
{
    static void Main()
    {
        // Path to the source FODS file (OpenDocument Flat XML Spreadsheet)
        string fodsFilePath = "input.fods";

        // Path where the resulting JSON will be saved
        string jsonOutputPath = "output.json";

        // -----------------------------------------------------------------
        // Load the FODS file into a Workbook using LoadOptions with the
        // appropriate LoadFormat (Fods). This respects the lifecycle rule
        // of creating and loading a workbook.
        // -----------------------------------------------------------------
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Fods);
        Workbook workbook = new Workbook(fodsFilePath, loadOptions);

        // -----------------------------------------------------------------
        // Configure JSON save options to preserve the full structure,
        // include empty cells, keep header rows, and format the output
        // with indentation for readability.
        // -----------------------------------------------------------------
        JsonSaveOptions jsonOptions = new JsonSaveOptions
        {
            ExportEmptyCells = true,          // Preserve empty cells as null
            HasHeaderRow = true,              // Treat the first row as header
            ExportNestedStructure = true,     // Keep parent‑child hierarchy
            SkipEmptyRows = false,            // Do not skip empty rows
            Indent = "  "                     // Two‑space indentation
        };

        // -----------------------------------------------------------------
        // Save the workbook as a JSON file using the configured options.
        // This follows the required save lifecycle rule.
        // -----------------------------------------------------------------
        workbook.Save(jsonOutputPath, jsonOptions);

        Console.WriteLine($"FODS file '{fodsFilePath}' has been converted to JSON at '{jsonOutputPath}'.");
    }
}