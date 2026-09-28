// Title: Export a specific Excel cell range to an indented JSON file using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads an .xlsx workbook with Aspose.Cells, extracts a defined cell area (e.g., A1:C10), and writes the data to a pretty‑printed JSON file. | Show how to combine Worksheet.Cells.ExportDataTable and System.Text.Json to serialize a selected range of cells into JSON, preserving column headers. | Create a reusable method ExportRangeToJson(string srcPath, string dstPath, string range) that validates the workbook, exports the range, and saves the JSON using Aspose.Cells.
// Common Searches: C# Aspose.Cells export selected Excel range to JSON file | How to use ExportDataTable with System.Text.Json for Excel to JSON conversion | Save worksheet cells A1:C10 as formatted JSON using Aspose.Cells .NET | Convert Excel cell area to JSON with column headers in C#
// Tags: excel range to json conversion Aspose.Cells | worksheet cells export json C# | pretty printed json from excel data | Aspose.Cells selected cells json output | convert cell area to json .NET

using System;
using System.Data;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

// The example loads "input.xlsx", extracts cells A1:C10 from the first worksheet using Aspose.Cells, converts the range to a DataTable, serializes it to indented JSON with System.Text.Json, and writes the result to "output.json".
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.json";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index or name as needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Define the cell range to export (e.g., A1:C10)
            CellArea exportRange = CellArea.CreateCellArea("A1", "C10");

            // Calculate range dimensions
            int startRow = exportRange.StartRow;
            int startColumn = exportRange.StartColumn;
            int rowCount = exportRange.EndRow - exportRange.StartRow + 1;
            int columnCount = exportRange.EndColumn - exportRange.StartColumn + 1;

            // Export the range to a DataTable (including column names)
            DataTable dataTable = worksheet.Cells.ExportDataTable(startRow, startColumn, rowCount, columnCount, true);

            // Convert the DataTable to JSON using System.Text.Json
            string jsonResult = JsonSerializer.Serialize(dataTable, new JsonSerializerOptions { WriteIndented = true });

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Write the JSON string to the output file
            File.WriteAllText(outputPath, jsonResult);
            Console.WriteLine($"Export completed successfully. JSON saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
