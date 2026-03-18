using System;
using Aspose.Cells;
using Aspose.Cells.Loading;
using Aspose.Cells.Saving;

class DbfToJsonConverter
{
    static void Main()
    {
        // Path to the source DBF file
        string dbfFilePath = "input.dbf";

        // Load the DBF workbook with default load options
        DbfLoadOptions loadOptions = new DbfLoadOptions();
        Workbook workbook = new Workbook(dbfFilePath, loadOptions);

        // Configure JSON save options to preserve data types and structure
        JsonSaveOptions jsonOptions = new JsonSaveOptions
        {
            // Keep original data types (numbers remain numbers, dates remain dates, etc.)
            ExportAsString = false,
            // Export the worksheet as a nested JSON structure (preserves hierarchy)
            ExportNestedStructure = true,
            // Skip rows that are completely empty
            SkipEmptyRows = true,
            // Treat the first row as header (column names)
            HasHeaderRow = true,
            // Optional: format the JSON with indentation for readability
            Indent = "  "
        };

        // Save the workbook as a JSON file
        string jsonOutputPath = "output.json";
        workbook.Save(jsonOutputPath, jsonOptions);

        Console.WriteLine($"DBF workbook has been successfully converted to JSON at: {jsonOutputPath}");
    }
}