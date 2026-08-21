// Title: C# – Access and Refresh the First Pivot Table on a Worksheet with Aspose.Cells
// Description: Load an existing workbook, select the first worksheet, retrieve its first PivotTable, call RefreshData and CalculateData, then save the updated file. Includes error handling for missing files and empty PivotTable collections.
// Keywords: Aspose.Cells | C# | .NET | PivotTable | first pivot table | refresh pivot data | recalculate pivot | worksheet pivot collection | load workbook | save workbook | example code
// Common Searches: Aspose.Cells get first pivot table C# | refresh pivot table data using Aspose.Cells | how to recalculate a pivot table in .NET | check for pivot tables on a worksheet Aspose.Cells | C# code to load workbook and refresh pivot
// Developer Intent: Retrieve the first PivotTable on the first worksheet, refresh its source data, recalculate, and persist the changes.
// Use Cases: Update pivot data after programmatically modifying the source range. | Ensure pivot totals are current before exporting or printing the workbook. | Validate the presence of a pivot table before running batch processing on multiple worksheets.
// AI Prompts: Generate C# code that accesses the second pivot table on the third worksheet and applies a filter to a specific field using Aspose.Cells. | Write a method to add a new pivot table to a worksheet, define its data source, and configure row and column fields with Aspose.Cells for .NET. | Provide example code that iterates through all pivot tables in a workbook, refreshes each one, and logs their names.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotDemo
{
    // Load an existing workbook, select the first worksheet, retrieve its first PivotTable, call RefreshData and CalculateData, then save the updated file. Includes error handling for missing files and empty PivotTable collections.
    public class AccessFirstPivotTable
    {
        public static void Run()
        {
            const string inputPath = "InputWithPivot.xlsx";
            const string outputPath = "OutputAfterPivotAccess.xlsx";

            // Verify input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            try
            {
                // Load the workbook containing at least one pivot table
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet (index 0)
                Worksheet firstWorksheet = workbook.Worksheets[0];

                // Get the collection of pivot tables on this worksheet
                PivotTableCollection pivotTables = firstWorksheet.PivotTables;

                if (pivotTables.Count > 0)
                {
                    // Retrieve the first pivot table (index 0)
                    PivotTable firstPivot = pivotTables[0];

                    // Refresh data for the pivot table
                    firstPivot.RefreshData();

                    // Recalculate the pivot table
                    firstPivot.CalculateData();

                    // Additional operations on 'firstPivot' can be placed here
                }
                else
                {
                    Console.WriteLine("No pivot tables found on the first worksheet.");
                }

                // Save the workbook after modifications
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred while processing the workbook: {ex.Message}");
            }
        }
    }

    public class Program
    {
        // Entry point required for compilation
        public static void Main(string[] args)
        {
            try
            {
                AccessFirstPivotTable.Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }
}
