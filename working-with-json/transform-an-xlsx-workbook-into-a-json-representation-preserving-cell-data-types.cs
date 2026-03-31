using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

class ExcelToJsonConverter
{
    static void Main()
    {
        // Path to the source Excel file
        string sourcePath = "input.xlsx";

        // Path where the JSON output will be saved
        string jsonOutputPath = "output.json";

        // Load the workbook from the file (lifecycle: load)
        Workbook workbook = new Workbook(sourcePath);

        // Configure JSON save options to preserve data types and full structure
        JsonSaveOptions jsonOptions = new JsonSaveOptions
        {
            // Export the Excel file as a JSON structure that mirrors the workbook
            ToExcelStruct = true,

            // Keep empty cells (null) so the layout is preserved
            ExportEmptyCells = true,

            // Export values in their native types (numbers, dates, booleans, etc.)
            ExportAsString = false,

            // Ensure the result is always a JSON object even for a single sheet
            AlwaysExportAsJsonObject = true
        };

        // Save the workbook as JSON (lifecycle: save)
        workbook.Save(jsonOutputPath, jsonOptions);

        Console.WriteLine($"Workbook has been converted to JSON and saved at: {jsonOutputPath}");
    }
}