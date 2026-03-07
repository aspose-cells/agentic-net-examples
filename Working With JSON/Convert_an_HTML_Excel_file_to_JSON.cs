using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsHtmlToJson
{
    class Program
    {
        static void Main()
        {
            // Path to the source HTML file (Excel saved as HTML)
            string htmlPath = "input.html";

            // Path where the resulting JSON will be saved
            string jsonPath = "output.json";

            // Load the HTML file into a Workbook.
            // HtmlLoadOptions can be customized if needed (e.g., TableToListObject).
            HtmlLoadOptions loadOptions = new HtmlLoadOptions();
            Workbook workbook = new Workbook(htmlPath, loadOptions);

            // Configure JSON save options.
            JsonSaveOptions jsonOptions = new JsonSaveOptions
            {
                // Export the workbook as a JSON object even if it contains a single sheet.
                AlwaysExportAsJsonObject = true,
                // Preserve the Excel structure (sheets, tables, etc.) in the JSON.
                ToExcelStruct = true,
                // Export empty cells and include header rows if present.
                ExportEmptyCells = true,
                HasHeaderRow = true
            };

            // Save the workbook as JSON.
            workbook.Save(jsonPath, jsonOptions);

            Console.WriteLine($"HTML workbook '{htmlPath}' has been converted to JSON at '{jsonPath}'.");
        }
    }
}