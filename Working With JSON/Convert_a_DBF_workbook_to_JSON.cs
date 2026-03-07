using System;
using Aspose.Cells;
using Aspose.Cells.Loading;
using Aspose.Cells.Saving;

class DbfToJsonConverter
{
    static void Main()
    {
        // Path to the source DBF file
        string dbfPath = "input.dbf";

        // Path for the resulting JSON file
        string jsonPath = "output.json";

        // Load the DBF file using DbfLoadOptions (default constructor)
        DbfLoadOptions loadOptions = new DbfLoadOptions();
        Workbook workbook = new Workbook(dbfPath, loadOptions);

        // Configure JSON save options
        JsonSaveOptions jsonOptions = new JsonSaveOptions
        {
            // Export cell values as strings (optional)
            ExportAsString = true,
            // Include empty cells as null (optional)
            ExportEmptyCells = true,
            // Treat the first row as header (optional)
            HasHeaderRow = true,
            // Indent the JSON for readability (optional)
            Indent = "  "
            // ExportArea is omitted to export the whole workbook
        };

        // Save the workbook as JSON using the configured options
        workbook.Save(jsonPath, jsonOptions);

        Console.WriteLine($"DBF file '{dbfPath}' has been converted to JSON at '{jsonPath}'.");
    }
}