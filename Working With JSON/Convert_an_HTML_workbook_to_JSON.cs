using System;
using Aspose.Cells;

class HtmlToJsonConverter
{
    static void Main()
    {
        // Input HTML workbook file path
        string htmlFilePath = "input.html";

        // Output JSON file path
        string jsonFilePath = "output.json";

        // Load the HTML file into a Workbook instance
        Workbook workbook = new Workbook(htmlFilePath);

        // Configure JSON export options
        JsonSaveOptions jsonOptions = new JsonSaveOptions
        {
            // Export empty cells as null (optional, adjust as needed)
            ExportEmptyCells = true,
            // Treat the first row as header (optional)
            HasHeaderRow = true,
            // Always export as a JSON object even if there is only one worksheet
            AlwaysExportAsJsonObject = true
        };

        // Save the workbook as a JSON file using the configured options
        workbook.Save(jsonFilePath, jsonOptions);
    }
}