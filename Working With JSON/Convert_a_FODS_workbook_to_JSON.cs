using System;
using Aspose.Cells;
using Aspose.Cells.Utility; // Required for JsonSaveOptions

public class ConvertFodsToJson
{
    public static void Main()
    {
        // Path to the source FODS file
        string sourcePath = "input.fods";

        // Path where the resulting JSON will be saved
        string outputPath = "output.json";

        // Load the FODS workbook with appropriate load options
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Fods);
        Workbook workbook = new Workbook(sourcePath, loadOptions);

        // Configure JSON save options
        JsonSaveOptions jsonOptions = new JsonSaveOptions
        {
            // Export as a JSON object even if there is only one worksheet
            AlwaysExportAsJsonObject = true,
            // Treat the first row as header names
            HasHeaderRow = true,
            // Include empty cells in the output
            ExportEmptyCells = true,
            // Do not create a nested parent‑child hierarchy
            ExportNestedStructure = false
        };

        // Save the workbook as JSON using the configured options
        workbook.Save(outputPath, jsonOptions);

        Console.WriteLine("FODS workbook successfully converted to JSON.");
    }
}