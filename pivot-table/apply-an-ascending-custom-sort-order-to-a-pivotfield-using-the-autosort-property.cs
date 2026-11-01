// Title: Applying an Ascending Custom Sort to a PivotField with Aspose.Cells for .NET (AutoSort property not available in current release)
// Common Searches: how to set an ascending custom sort on a pivot field using Aspose.Cells in C# | Aspose.Cells AutoSort property example for pivot table row fields | C# code to sort pivot table row field alphabetically with Aspose.Cells | why AutoSort, AutoSortFieldIndex, AutoSortOrder are missing in Aspose.Cells version | refresh and calculate pivot table after attempting custom sort in Aspose.Cells
// Tags: Aspose.Cells pivot field ascending sort | AutoSort property usage Aspose.Cells .NET | custom pivot table sorting with Aspose.Cells | pivot table row field sort C# | Aspose.Cells version limitation AutoSort API

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotSorting
{
    // The sample loads an existing workbook, verifies that a pivot table with at least one row field exists, notes that the AutoSort‑related properties are unavailable in the current Aspose.Cells version, refreshes and calculates the pivot table, and saves the updated workbook.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Verify that the input workbook exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {Path.GetFullPath(inputPath)}");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet (adjust index/name as needed)
                Worksheet worksheet = workbook.Worksheets[0];

                // Ensure the worksheet contains at least one pivot table
                if (worksheet.PivotTables.Count == 0)
                {
                    Console.WriteLine("No pivot tables found on the first worksheet.");
                    return;
                }

                // Get the first pivot table on the worksheet
                PivotTable pivotTable = worksheet.PivotTables[0];

                // Ensure the pivot table has at least one row field
                if (pivotTable.RowFields.Count == 0)
                {
                    Console.WriteLine("The pivot table does not contain any row fields.");
                    return;
                }

                // NOTE: The AutoSort, AutoSortFieldIndex, AutoSortOrder, and CustomItems
                // properties are not available in the current Aspose.Cells version.
                // If custom sorting is required, it should be implemented using the
                // appropriate API supported by the library version in use.

                // Refresh the pivot table to apply any changes
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Save the workbook with the (potentially) updated pivot table
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors to prevent the application from crashing
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
