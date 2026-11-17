// Title: C# console app that prints the first worksheet of an Excel file as formatted JSON using Aspose.Cells
// Common Searches: how to read the first sheet of an Excel workbook and output JSON in a C# console program | Aspose.Cells export first worksheet to JSON string in .NET Core console application | convert Excel sheet to JSON with System.Text.Json and Aspose.Cells in C# | C# command line tool for converting the first worksheet of an .xlsx file to indented JSON | example code for printing worksheet data as JSON using Aspose.Cells library
// Tags: Aspose.Cells first worksheet JSON export | C# console application Excel JSON output | System.Text.Json DataTable conversion | determine used range Aspose.Cells worksheet | standard output JSON printing C#

using System;
using System.Data;
using System.IO;
using Aspose.Cells;
using System.Text.Json;

namespace WorksheetJsonExporter
{
    // A C# console program that accepts an Excel file path, loads the workbook with Aspose.Cells, extracts the first worksheet's used range into a DataTable, serializes the table to indented JSON using System.Text.Json, and writes the JSON to standard output, with error handling for missing files and runtime exceptions.
    class Program
    {
        static void Main(string[] args)
        {
            // Expect the first argument to be the path of the Excel file.
            if (args.Length == 0)
            {
                Console.Error.WriteLine("Usage: WorksheetJsonExporter <excel-file-path>");
                return;
            }

            string filePath = args[0];

            // Verify that the file exists before attempting to load it.
            if (!File.Exists(filePath))
            {
                Console.Error.WriteLine($"Error: File not found - {filePath}");
                return;
            }

            try
            {
                // Load the workbook from the specified file.
                Workbook workbook = new Workbook(filePath);

                // Get the first worksheet (index 0).
                Worksheet sheet = workbook.Worksheets[0];

                // Determine the used range of the worksheet.
                int maxRow = sheet.Cells.MaxDataRow;
                int maxColumn = sheet.Cells.MaxDataColumn;

                // Export the used range to a DataTable.
                // The first row is treated as column names (true).
                DataTable dt = sheet.Cells.ExportDataTable(
                    0,                 // start row
                    0,                 // start column
                    maxRow + 1,        // total rows
                    maxColumn + 1,     // total columns
                    true);             // export column names

                // Convert the DataTable to JSON using System.Text.Json.
                var options = new JsonSerializerOptions { WriteIndented = true };
                string json = JsonSerializer.Serialize(dt, options);

                // Output the JSON to standard output.
                Console.WriteLine(json);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
