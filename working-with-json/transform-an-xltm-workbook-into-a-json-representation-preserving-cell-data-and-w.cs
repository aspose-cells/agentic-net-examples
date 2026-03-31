using System;
using Aspose.Cells;

class XltmToJsonConverter
{
    // Converts an XLTM workbook to a JSON file while preserving the Excel structure.
    public static void Convert(string xltmPath, string jsonOutputPath)
    {
        // Load the XLTM workbook from the specified file path.
        Workbook workbook = new Workbook(xltmPath);

        // Configure JSON save options to keep the workbook's structural information.
        JsonSaveOptions saveOptions = new JsonSaveOptions
        {
            // Preserve the Excel hierarchy (worksheets, cells, styles, etc.).
            ToExcelStruct = true,

            // Ensure the output is a JSON object even if there is only one worksheet.
            AlwaysExportAsJsonObject = true,

            // Export empty cells as null to keep the cell layout intact.
            ExportEmptyCells = true,

            // Do not treat the first row as a header row; export all rows as data.
            HasHeaderRow = false
        };

        // Save the workbook as a JSON file using the configured options.
        workbook.Save(jsonOutputPath, saveOptions);
    }

    static void Main()
    {
        // Example usage: convert "template.xltm" to "result.json".
        string inputPath = "template.xltm";
        string outputPath = "result.json";

        Convert(inputPath, outputPath);
        Console.WriteLine($"XLTM workbook '{inputPath}' has been converted to JSON at '{outputPath}'.");
    }
}