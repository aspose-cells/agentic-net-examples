using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

class CsvToJson
{
    static void Main()
    {
        // Path to the tab‑delimited CSV file
        string csvPath = "data.tsv";

        // Create a new workbook
        Workbook workbook = new Workbook();

        // Import the CSV data into the first worksheet.
        // "\t" specifies tab delimiter, true enables numeric conversion.
        workbook.Worksheets[0].Cells.ImportCSV(csvPath, "\t", true, 0, 0);

        // Set JSON export options.
        JsonSaveOptions jsonOptions = new JsonSaveOptions
        {
            ExportNestedStructure = true,   // Preserve hierarchical structure if applicable
            Indent = "    "                 // Pretty‑print with 4‑space indentation
        };

        // Save the workbook as a JSON file using the configured options.
        workbook.Save("output.json", jsonOptions);
    }
}