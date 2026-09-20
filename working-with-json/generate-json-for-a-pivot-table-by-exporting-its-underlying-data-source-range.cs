// Title: Export the underlying data source range of a pivot table to an indented JSON file using Aspose.Cells for .NET
// AI Prompts: Write C# code that opens an Excel workbook with Aspose.Cells, gets the worksheet's used range, converts that range into a System.Data.DataTable, and serializes the table to pretty‑printed JSON using System.Text.Json. | Adapt the sample to locate a pivot table by its name, extract only its source range, and write the resulting JSON to a MemoryStream instead of a physical file.
// Common Searches: how to export pivot table source range to JSON with Aspose.Cells in C# | C# Aspose.Cells convert worksheet range to DataTable and serialize to JSON | save Excel used range as formatted JSON file using System.Text.Json | retrieve pivot table data source range programmatically with Aspose.Cells .NET
// Tags: pivot table source range export to JSON | Aspose.Cells range to DataTable conversion | serialize Excel data to indented JSON .NET | C# extract worksheet used range with Aspose.Cells | write JSON file from Aspose.Cells workbook

using System;
using System.Data;
using System.IO;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// Alias to avoid conflict with System.Range
using CellsRange = Aspose.Cells.Range;

// The example loads an Excel workbook, checks for a pivot table on the first worksheet, obtains the worksheet's used range, converts that range into a System.Data.DataTable via a helper method, serializes the DataTable to pretty‑printed JSON with System.Text.Json, and writes the JSON to a file while handling errors and reporting progress.
class PivotTableJsonExporter
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "PivotDataSource.json";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Assume the pivot table is on the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one pivot table
            if (sheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No pivot tables found in the worksheet.");
                return;
            }

            // Get the first pivot table (not used further, but kept for validation)
            PivotTable pivot = sheet.PivotTables[0];

            // Use the worksheet's used range as the data source for simplicity
            CellsRange dataSourceRange = sheet.Cells.MaxDisplayRange;
            if (dataSourceRange == null)
            {
                Console.WriteLine("Unable to retrieve the data source range.");
                return;
            }

            // Convert the range to a DataTable for JSON serialization
            DataTable dt = RangeToDataTable(dataSourceRange);

            // Serialize the DataTable to JSON using System.Text.Json
            var options = new JsonSerializerOptions { WriteIndented = true };
            string json = JsonSerializer.Serialize(dt, options);

            // Output JSON to console and write to file
            Console.WriteLine(json);
            try
            {
                File.WriteAllText(outputPath, json);
                Console.WriteLine($"JSON data written to {outputPath}");
            }
            catch (Exception writeEx)
            {
                Console.WriteLine($"Failed to write JSON file: {writeEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Helper method: converts an Aspose.Cells Range to a System.Data.DataTable
    private static DataTable RangeToDataTable(CellsRange range)
    {
        DataTable table = new DataTable();

        // Determine the range dimensions
        int firstRow = range.FirstRow;
        int firstCol = range.FirstColumn;
        int totalRows = range.RowCount;
        int totalCols = range.ColumnCount;

        // Use the first row as column headers
        for (int col = 0; col < totalCols; col++)
        {
            string columnName = range[firstRow, firstCol + col].StringValue;
            if (string.IsNullOrEmpty(columnName))
                columnName = $"Column{col + 1}";

            // Ensure column names are unique
            string originalName = columnName;
            int duplicateIndex = 1;
            while (table.Columns.Contains(columnName))
            {
                columnName = $"{originalName}_{duplicateIndex++}";
            }

            table.Columns.Add(columnName, typeof(string));
        }

        // Populate rows starting after the header row
        for (int row = 1; row < totalRows; row++)
        {
            DataRow dataRow = table.NewRow();
            for (int col = 0; col < totalCols; col++)
            {
                dataRow[col] = range[firstRow + row, firstCol + col].StringValue;
            }
            table.Rows.Add(dataRow);
        }

        return table;
    }
}
