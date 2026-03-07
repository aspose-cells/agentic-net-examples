using System;
using Aspose.Cells;

class ExcelToJsonConverter
{
    // Converts an XLSX file to a JSON file using Aspose.Cells
    public static void ConvertXlsxToJson(string sourcePath, string jsonPath)
    {
        // Load the Excel workbook from the specified file
        Workbook workbook = new Workbook(sourcePath);

        // Configure JSON save options
        JsonSaveOptions saveOptions = new JsonSaveOptions
        {
            // Export as a JSON object even if the workbook contains a single worksheet
            AlwaysExportAsJsonObject = true,
            // Include empty cells in the output
            ExportEmptyCells = true,
            // Treat the first row as a header row
            HasHeaderRow = true
        };

        // Save the workbook as a JSON file using the configured options
        workbook.Save(jsonPath, saveOptions);
    }

    static void Main()
    {
        // Example file paths (adjust as needed)
        string sourceFile = "input.xlsx";
        string jsonFile = "output.json";

        // Perform the conversion
        ConvertXlsxToJson(sourceFile, jsonFile);

        Console.WriteLine($"Conversion completed: '{sourceFile}' -> '{jsonFile}'");
    }
}