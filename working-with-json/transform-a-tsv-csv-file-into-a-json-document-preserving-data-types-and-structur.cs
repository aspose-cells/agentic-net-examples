using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsJsonExport
{
    class Program
    {
        static void Main()
        {
            // Path to the source TSV (tab‑separated) file
            string tsvPath = "data.tsv";

            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Import the TSV data.
            // Splitter is a tab character, numeric conversion is enabled,
            // and data is placed starting at cell A1 (row 0, column 0).
            cells.ImportCSV(tsvPath, "\t", true, 0, 0);

            // Determine the range that contains the imported data.
            // MaxDisplayRange returns the smallest range that includes all non‑empty cells.
            Aspose.Cells.Range dataRange = cells.MaxDisplayRange;

            // Configure JSON export options.
            // ExportNestedStructure = true keeps the original hierarchy (rows/columns).
            // ExportAsString is left false (default) so numeric and date types are preserved.
            JsonSaveOptions jsonOptions = new JsonSaveOptions
            {
                ExportNestedStructure = true,
                ExportEmptyCells = true,
                HasHeaderRow = true
            };

            // Export the range to a JSON string.
            string jsonOutput = JsonUtility.ExportRangeToJson(dataRange, jsonOptions);

            // Write the JSON string to a file.
            string jsonPath = "output.json";
            File.WriteAllText(jsonPath, jsonOutput);

            // (Optional) Save the workbook as an Excel file for verification.
            workbook.Save("output.xlsx", SaveFormat.Xlsx);

            Console.WriteLine($"TSV data has been exported to JSON file: {jsonPath}");
        }
    }
}