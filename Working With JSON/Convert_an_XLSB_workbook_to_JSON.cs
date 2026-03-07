using System;
using Aspose.Cells;

class XlsbToJsonConverter
{
    // Converts an XLSB workbook to a JSON file.
    public static void ConvertXlsbToJson(string xlsbPath, string jsonPath)
    {
        // Load the XLSB workbook from the specified file.
        Workbook workbook = new Workbook(xlsbPath);

        // Configure JSON save options.
        JsonSaveOptions jsonOptions = new JsonSaveOptions();
        jsonOptions.AlwaysExportAsJsonObject = true;   // Export as a JSON object even if there is only one worksheet.
        jsonOptions.ExportNestedStructure = true;      // Preserve the hierarchical structure of the workbook.
        jsonOptions.SkipEmptyRows = true;              // Omit empty rows from the output.

        // Save the workbook as a JSON file using the configured options.
        workbook.Save(jsonPath, jsonOptions);
    }

    static void Main()
    {
        string sourceXlsb = "input.xlsb";   // Path to the source XLSB file.
        string targetJson = "output.json";  // Desired path for the resulting JSON file.

        ConvertXlsbToJson(sourceXlsb, targetJson);
        Console.WriteLine("XLSB to JSON conversion completed successfully.");
    }
}