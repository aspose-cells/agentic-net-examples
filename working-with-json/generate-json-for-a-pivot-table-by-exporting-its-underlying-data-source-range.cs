// Title: Export a PivotTable’s source data range to an indented JSON file using Aspose.Cells for .NET
// AI Prompts: Write C# code that opens an Excel workbook, locates the first PivotTable, extracts its data source range, converts the range to a DataTable, and saves the DataTable as a pretty‑printed JSON file with Aspose.Cells. | Modify the example to accept a worksheet name and pivot table index as parameters, and output the JSON to a MemoryStream instead of a physical file. | Add robust error handling that logs missing files, absent PivotTables, empty source ranges, and saves the workbook with a timestamped filename while preserving the original.
// Common Searches: c# aspose.cells export pivot table source range to json file | retrieve pivot table data source address using Aspose.Cells .NET | convert excel range to datatable and serialize to formatted json in C# | save workbook after exporting pivot source data with Aspose.Cells | asp.net example for exporting pivot source data to json
// Tags: aspose.cells export pivot source range json | pivot table data source address retrieval c# | excel range to datatable conversion aspnet | pretty printed json serialization datatable | timestamped workbook save aspose.cells

using System;
using System.Data;
using System.IO;
using System.Linq;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsPivotExport
{
    // The program loads an Excel workbook, obtains the address of the first PivotTable's data source range, creates a range object, exports that range to a DataTable with column headers, serializes the DataTable to an indented JSON file, writes the JSON to disk, and finally saves a copy of the workbook.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputFile = "InputWithPivot.xlsx";
            const string jsonOutputFile = "PivotSourceData.json";
            const string savedWorkbookFile = "InputWithPivot_Saved.xlsx";

            try
            {
                // Verify that the input workbook exists
                if (!File.Exists(inputFile))
                {
                    Console.WriteLine($"Error: The file '{inputFile}' was not found.");
                    return;
                }

                // Load the workbook that contains the pivot table
                Workbook workbook = new Workbook(inputFile);

                // Get the first worksheet (adjust index if needed)
                Worksheet sheet = workbook.Worksheets[0];

                // Ensure there is at least one pivot table
                if (sheet.PivotTables.Count == 0)
                {
                    Console.WriteLine("Error: No pivot tables found on the first worksheet.");
                    return;
                }

                // Access the first pivot table
                PivotTable pivot = sheet.PivotTables[0];

                // Retrieve the address of the data source range used by the pivot table
                string sourceRangeAddress = pivot.DataSource?.FirstOrDefault();
                if (string.IsNullOrEmpty(sourceRangeAddress))
                {
                    Console.WriteLine("Error: Pivot table data source address is unavailable.");
                    return;
                }

                // Create a range object for the source data (use alias to avoid ambiguity)
                AsposeRange sourceRange = sheet.Cells.CreateRange(sourceRangeAddress);

                // Export the range to a DataTable (first row as column names)
                DataTable dt = sheet.Cells.ExportDataTable(
                    sourceRange.FirstRow,
                    sourceRange.FirstColumn,
                    sourceRange.RowCount,
                    sourceRange.ColumnCount,
                    true);

                // Serialize the DataTable to JSON
                string json = JsonSerializer.Serialize(dt, new JsonSerializerOptions { WriteIndented = true });

                // Write JSON to file
                File.WriteAllText(jsonOutputFile, json);
                Console.WriteLine($"Pivot source data exported to '{jsonOutputFile}'.");

                // Save the workbook (demonstrating the required save rule)
                workbook.Save(savedWorkbookFile);
                Console.WriteLine($"Workbook saved as '{savedWorkbookFile}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
