// Title: Export Excel worksheet range to an indented JSON file with column headers using Aspose.Cells in C#
// Common Searches: how to convert an Excel worksheet to a pretty‑printed JSON file using Aspose.Cells C# | Aspose.Cells export DataTable to JSON with indentation and column names | C# write Excel data to formatted JSON using System.Text.Json and Aspose.Cells | save Excel table as readable JSON preserving headers in .NET | export used range of an Excel sheet to indented JSON in C#
// Tags: Aspose.Cells export worksheet to JSON | C# serialize DataTable with System.Text.Json | Excel used range to formatted JSON | create output folder for JSON export | write indented JSON file from Excel data

using System;
using System.Data;
using System.IO;
using Aspose.Cells;
using System.Text.Json;

// The example loads an Excel workbook with Aspose.Cells, extracts the used range of the first worksheet into a DataTable (using the first row as column headers), serializes the DataTable to a pretty‑printed JSON string via System.Text.Json, ensures the target directory exists, and writes the JSON to a file while handling missing‑file and generic exceptions.
class ExportTableToJson
{
    static void Main()
    {
        // Paths for input Excel and output JSON
        string excelPath = @"C:\Input\Sample.xlsx";
        string jsonPath = @"C:\Output\Result.json";

        try
        {
            // Verify that the source Excel file exists
            if (!File.Exists(excelPath))
            {
                Console.WriteLine($"Error: Input file not found at '{excelPath}'.");
                return;
            }

            // Load the workbook (lifecycle rule: load)
            Workbook workbook = new Workbook(excelPath);

            // Get the first worksheet (adjust index or name as needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Determine the used range of the worksheet
            Aspose.Cells.Range usedRange = sheet.Cells.MaxDisplayRange;

            // Export the used range to a DataTable.
            // The first row of the range will be used as column headers.
            DataTable table = sheet.Cells.ExportDataTable(
                usedRange.FirstRow,          // start row
                usedRange.FirstColumn,       // start column
                usedRange.RowCount,          // total rows
                usedRange.ColumnCount,       // total columns
                true);                       // use first row as column names

            // Serialize the DataTable to JSON with indentation for readability
            var jsonOptions = new JsonSerializerOptions { WriteIndented = true };
            string json = JsonSerializer.Serialize(table, jsonOptions);

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(jsonPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Write the JSON string to a file (lifecycle rule: save)
            File.WriteAllText(jsonPath, json);

            Console.WriteLine("Export completed successfully.");
        }
        catch (FileNotFoundException fnfEx)
        {
            Console.WriteLine($"File not found: {fnfEx.Message}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
