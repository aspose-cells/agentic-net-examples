using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Path to the source XLSX file
        string sourcePath = "input.xlsx";

        // Path where the JSON output will be saved
        string jsonPath = "output.json";

        // Load the Excel workbook from the XLSX file
        Workbook workbook = new Workbook(sourcePath);

        // Configure JSON save options (customize as needed)
        JsonSaveOptions jsonOptions = new JsonSaveOptions();
        // Export as a JSON object even if there is only one worksheet
        jsonOptions.AlwaysExportAsJsonObject = true;
        // Include empty cells in the JSON output (null values)
        jsonOptions.ExportEmptyCells = true;
        // Treat the first row as header if present
        jsonOptions.HasHeaderRow = true;

        // Save the workbook as a JSON file using the configured options
        workbook.Save(jsonPath, jsonOptions);

        Console.WriteLine("Excel workbook has been successfully converted to JSON.");
    }
}