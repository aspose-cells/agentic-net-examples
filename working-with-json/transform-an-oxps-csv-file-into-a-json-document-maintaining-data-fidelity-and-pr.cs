using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

class CsvToJsonConverter
{
    static void Main()
    {
        // Path to the source CSV file (OXPS CSV)
        string csvPath = "input.csv";

        // Desired output JSON file path
        string jsonPath = "output.json";

        // Load the CSV file into a workbook using CSV load options
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Csv);
        Workbook workbook = new Workbook(csvPath, loadOptions);

        // Configure JSON save options to preserve data fidelity and formatting
        JsonSaveOptions jsonOptions = new JsonSaveOptions
        {
            ExportEmptyCells = true,      // Export empty cells as null
            HasHeaderRow = true,          // Treat the first row as header
            ExportNestedStructure = false,// Flat structure (table-like)
            ExportAsString = false,       // Preserve original data types
            SkipEmptyRows = false,        // Keep empty rows
            Indent = "  "                 // Pretty‑print with indentation
        };

        // Save the workbook as a JSON document using the configured options
        workbook.Save(jsonPath, jsonOptions);

        Console.WriteLine("CSV file has been successfully converted to JSON.");
    }
}