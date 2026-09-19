// Title: Export a specific Excel table range to an indented JSON file with column headers using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an .xlsx workbook with Aspose.Cells, extracts a defined CellArea into a DataTable using ExportDataTable (including headers), and saves the DataTable as pretty‑printed JSON. | Show how to use System.Text.Json to serialize a DataTable with the WriteIndented option and write the resulting JSON to a file path.
// Common Searches: C# Aspose.Cells export selected worksheet range to JSON with column names as keys | How to convert an Excel table to formatted JSON file using Aspose.Cells and System.Text.Json | Save Excel range as indented JSON file preserving header row in .NET
// Tags: Aspose.Cells ExportDataTable to JSON | C# serialize DataTable with System.Text.Json | Excel range to indented JSON file | Aspose.Cells extract table with headers | pretty‑printed JSON output from Excel

using System;
using System.Data;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

// The example loads 'input.xlsx', extracts cells A1:D10 (including the header row) into a DataTable via Aspose.Cells' ExportDataTable, serializes the DataTable to indented JSON using System.Text.Json, and writes the result to 'output.json'.
class ExportTableToJson
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.json";

            // Verify that the input workbook exists to avoid FileNotFoundException.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the workbook that contains the table.
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index if needed).
            Worksheet sheet = workbook.Worksheets[0];

            // Define the area of the table to export (including the header row).
            // Example: A1:D10 => rows 0‑9, columns 0‑3
            CellArea tableArea = new CellArea
            {
                StartRow = 0,
                StartColumn = 0,
                EndRow = 9,
                EndColumn = 3
            };

            // Export the defined range to a DataTable (first row used as column headers).
            int totalRows = tableArea.EndRow - tableArea.StartRow + 1;
            int totalColumns = tableArea.EndColumn - tableArea.StartColumn + 1;
            DataTable dataTable = sheet.Cells.ExportDataTable(
                tableArea.StartRow,
                tableArea.StartColumn,
                totalRows,
                totalColumns,
                true);

            // Serialize the DataTable to formatted JSON.
            string json = JsonSerializer.Serialize(dataTable, new JsonSerializerOptions { WriteIndented = true });

            // Write JSON to the output file.
            File.WriteAllText(outputPath, json);
            Console.WriteLine($"JSON file has been saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message.
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
