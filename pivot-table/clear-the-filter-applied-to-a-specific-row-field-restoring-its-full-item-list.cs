// Title: How to clear a filter on a specific row field of an Aspose.Cells pivot table using C#
// Common Searches: Aspose.Cells C# clear filter for a pivot table row field | remove pivot row field filter programmatically with Aspose.Cells .NET | reset items in a pivot row field using the ClearFilter method | C# code example to clear a pivot table row field filter in Excel | using Aspose.Cells to delete filters from a pivot table row field
// Tags: Aspose.Cells ClearFilter PivotField | C# remove pivot row field filter | Aspose.Cells pivot table filter reset | Excel workbook pivot filter manipulation .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsRunner
{
    // The example loads an Excel workbook, accesses the first worksheet, retrieves a pivot table named "PivotTable1", selects a row field by its index, calls ClearFilter() on that PivotField to remove any applied filter, and saves the updated workbook.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "output.xlsx";

                // Verify that the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook containing the pivot table
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet (adjust index if needed)
                Worksheet sheet = workbook.Worksheets[0];

                // Get the pivot table collection from the worksheet
                PivotTableCollection pivotTables = sheet.PivotTables;

                // Locate the pivot table by its name
                string pivotName = "PivotTable1";

                // Retrieve the pivot table using its name
                PivotTable pivotTable = null;
                try
                {
                    pivotTable = pivotTables[pivotName];
                }
                catch (Exception)
                {
                    // Indexer throws if name not found; handle below
                }

                if (pivotTable == null)
                {
                    Console.WriteLine($"Pivot table '{pivotName}' not found.");
                    return;
                }

                // Identify the row field to clear the filter from (by index)
                int rowFieldIndex = 0; // change to the appropriate index
                if (rowFieldIndex < 0 || rowFieldIndex >= pivotTable.RowFields.Count)
                {
                    Console.WriteLine("Row field index is out of range.");
                    return;
                }

                // Clear any filter applied to this row field, restoring the full item list
                PivotField rowField = pivotTable.RowFields[rowFieldIndex];
                rowField.ClearFilter();

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook with the changes
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
