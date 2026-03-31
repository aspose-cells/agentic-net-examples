using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

class Program
{
    static void Main()
    {
        // Path to the source DBF file
        string sourcePath = "input.dbf";

        // Path where the resulting JSON will be saved
        string outputPath = "output.json";

        // Load the DBF file into a Workbook instance
        Workbook workbook = new Workbook(sourcePath);

        // Configure JSON save options to preserve data types and hierarchical structure
        JsonSaveOptions jsonOptions = new JsonSaveOptions
        {
            // Convert the workbook to a JSON structure that mirrors the Excel layout
            ToExcelStruct = true,

            // Export parent‑child hierarchy (nested JSON)
            ExportNestedStructure = true,

            // Keep original data types (numbers, dates, etc.) instead of forcing strings
            ExportAsString = false,

            // Treat the first row as a header row if it exists
            HasHeaderRow = true,

            // Do not skip empty rows so the original layout is retained
            SkipEmptyRows = false,

            // Ensure the output is a JSON object even when there is only one worksheet
            AlwaysExportAsJsonObject = true,

            // Add indentation for readability
            Indent = "  "
        };

        // Save the workbook as a JSON file using the configured options
        workbook.Save(outputPath, jsonOptions);

        // Optional: display the generated JSON content
        Console.WriteLine(File.ReadAllText(outputPath));
    }
}