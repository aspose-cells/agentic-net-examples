using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsTsvToJson
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source TSV file
            string tsvPath = "input.tsv";

            // Path for the output JSON file
            string jsonPath = "output.json";

            // Create a new workbook (in-memory Excel file)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Import the TSV data into the worksheet.
            // - splitter: "\t" for tab delimiter
            // - convertNumericData: true to convert numbers automatically
            // - firstRow / firstColumn: start at cell A1 (0,0)
            cells.ImportCSV(tsvPath, "\t", true, 0, 0);

            // Configure JSON save options.
            JsonSaveOptions saveOptions = new JsonSaveOptions
            {
                // Export as a JSON object even if there is only one worksheet
                AlwaysExportAsJsonObject = true,
                // Assume the first row contains column headers
                HasHeaderRow = true,
                // Export empty cells as null (optional, can be set to false)
                ExportEmptyCells = true,
                // Export values as strings (optional)
                ExportAsString = false
            };

            // Save the workbook as a JSON file using the configured options
            workbook.Save(jsonPath, saveOptions);

            Console.WriteLine($"TSV file '{tsvPath}' has been converted to JSON file '{jsonPath}'.");
        }
    }
}