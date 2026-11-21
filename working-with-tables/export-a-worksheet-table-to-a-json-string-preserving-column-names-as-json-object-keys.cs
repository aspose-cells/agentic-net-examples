// Title: Export an Excel worksheet’s table to a formatted JSON string with column headers using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells export worksheet data to JSON preserving column names | How to convert an Excel table to a JSON string with headers using Aspose.Cells | Aspose.Cells ExportDataTable to JSON example in .NET | Serialize Excel worksheet range to indented JSON with System.Text.Json in C# | Export used range of first worksheet to JSON using Aspose.Cells API
// Tags: Aspose.Cells ExportDataTable to JSON | Excel worksheet to JSON serialization C# | preserve column headers JSON export Aspose.Cells | System.Text.Json serialize DataTable example | export used range worksheet Aspose.Cells

using System;
using System.Data;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

namespace ExportWorksheetToJsonApp
{
    // The sample loads an Excel workbook, extracts the used range of the first worksheet into a DataTable while keeping column names, then serializes the DataTable to an indented JSON string using System.Text.Json and writes the result to the console.
    class ExportWorksheetToJson
    {
        static void Main()
        {
            const string inputPath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Determine the used range of the worksheet
                int firstRow = worksheet.Cells.MinRow;
                int firstColumn = worksheet.Cells.MinColumn;
                int totalRows = worksheet.Cells.MaxRow - worksheet.Cells.MinRow + 1;
                int totalColumns = worksheet.Cells.MaxColumn - worksheet.Cells.MinColumn + 1;

                // Export the used range to a DataTable, preserving column names as keys
                DataTable dataTable = worksheet.Cells.ExportDataTable(
                    firstRow,
                    firstColumn,
                    totalRows,
                    totalColumns,
                    true // export column names
                );

                // Serialize the DataTable to a JSON string using System.Text.Json
                var jsonOptions = new JsonSerializerOptions { WriteIndented = true };
                string jsonString = JsonSerializer.Serialize(dataTable, jsonOptions);

                // Output the JSON string
                Console.WriteLine(jsonString);
            }
            catch (Exception ex)
            {
                // Handle any runtime errors gracefully
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
