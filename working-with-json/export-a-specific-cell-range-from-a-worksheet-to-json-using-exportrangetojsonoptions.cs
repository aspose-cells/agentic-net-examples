// Title: Export a specific cell range from an Excel worksheet to a JSON file using Aspose.Cells in C#
// Common Searches: Aspose.Cells C# export selected range to JSON file | How to convert Excel range A1:C5 to JSON with Aspose.Cells | Export worksheet cells with column headers to JSON in .NET | Aspose.Cells ExportRangeToJsonOptions example code | C# generate JSON from Excel range using Aspose.Cells library
// Tags: Aspose.Cells ExportRangeToJsonOptions usage | C# export Excel cell area to JSON | Excel range to JSON with column headers | Aspose.Cells CellArea JSON serialization | Generate JSON file from worksheet range .NET

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

// The sample loads or creates an Excel workbook, selects a cell area (e.g., A1:C5) on the first worksheet, uses Aspose.Cells ExportRangeToJsonOptions to convert the range into a formatted JSON string (optionally using the first row as column headers), and writes the JSON to an output file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.json";

            // Load workbook if it exists; otherwise create a simple workbook.
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                Worksheet ws = workbook.Worksheets[0];
                ws.Cells["A1"].PutValue("Header1");
                ws.Cells["B1"].PutValue("Header2");
                ws.Cells["C1"].PutValue("Header3");
                ws.Cells["A2"].PutValue("Data1");
                ws.Cells["B2"].PutValue(123);
                ws.Cells["C2"].PutValue(DateTime.Now);
                workbook.Save(inputPath);
            }

            // Get the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            // Define the cell range to export (A1:C5).
            CellArea range = CellArea.CreateCellArea("A1", "C5");

            // Export the range to a JSON string.
            string json = ExportRangeToJson(sheet, range, includeColumnNames: true);

            // Write JSON to file.
            File.WriteAllText(outputPath, json);
            Console.WriteLine($"Export completed successfully. JSON saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    /// <param name="sheet">Worksheet containing the range.</param>
    /// <param name="range">Cell area to export.</param>
    /// <param name="includeColumnNames">If true, first row is treated as column headers.</param>
    /// <returns>JSON representation of the range.</returns>
    private static string ExportRangeToJson(Worksheet sheet, CellArea range, bool includeColumnNames)
    {
        var cells = sheet.Cells;
        int startRow = range.StartRow;
        int endRow = range.EndRow;
        int startCol = range.StartColumn;
        int endCol = range.EndColumn;

        List<Dictionary<string, object>> rows = new List<Dictionary<string, object>>();
        List<string> headers = new List<string>();

        // If column names are required, read the first row as headers.
        if (includeColumnNames)
        {
            for (int col = startCol; col <= endCol; col++)
            {
                var headerValue = cells[startRow, col].Value?.ToString() ?? $"Column{col - startCol + 1}";
                headers.Add(headerValue);
            }
            startRow++; // Data starts from the next row.
        }
        else
        {
            // Generate default headers if not provided.
            for (int col = startCol; col <= endCol; col++)
            {
                headers.Add($"Column{col - startCol + 1}");
            }
        }

        // Iterate through each data row.
        for (int row = startRow; row <= endRow; row++)
        {
            var rowDict = new Dictionary<string, object>();
            bool isEmptyRow = true;

            for (int col = startCol; col <= endCol; col++)
            {
                var cell = cells[row, col];
                object value = cell.Value;

                // Treat empty cells as null.
                if (value == null)
                {
                    value = null;
                }
                else
                {
                    isEmptyRow = false;
                }

                rowDict[headers[col - startCol]] = value;
            }

            // Skip completely empty rows.
            if (!isEmptyRow)
            {
                rows.Add(rowDict);
            }
        }

        // Serialize the list of rows to JSON.
        var jsonOptions = new JsonSerializerOptions { WriteIndented = true };
        return JsonSerializer.Serialize(rows, jsonOptions);
    }
}
