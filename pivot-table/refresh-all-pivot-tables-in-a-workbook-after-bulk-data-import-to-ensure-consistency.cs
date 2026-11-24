// Title: Refresh every pivot table in an Excel workbook after bulk data import using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# refresh all pivot tables after updating worksheet data | How to programmatically refresh pivot tables in a workbook with Aspose.Cells .NET | Refresh pivot cache and recalculate values for multiple sheets using Aspose.Cells | C# code to iterate worksheets and refresh pivot tables after bulk import | Update pivot tables automatically after loading data into Excel with Aspose.Cells
// Tags: Aspose.Cells refresh pivot tables | pivot table cache refresh C# | calculate pivot data Aspose.Cells | bulk data import Excel Aspose.Cells | iterate worksheets pivot tables C#

using Aspose.Cells;
using Aspose.Cells.Pivot;
using System;
using System.IO;

// The example loads an existing Excel workbook, provides a placeholder for bulk data import, then loops through each worksheet and its pivot tables, calling RefreshData() to update the cache and CalculateData() to recompute values before saving the refreshed workbook. It includes file existence checks and exception handling.
class Program
{
    static void Main()
    {
        const string inputPath = "DataWorkbook.xlsx";
        const string outputPath = "DataWorkbook_Refreshed.xlsx";

        // Verify that the input workbook exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook that contains the data and pivot tables
            Workbook workbook = new Workbook(inputPath);

            // --------------------------------------------------------------------
            // Bulk data import would occur here.
            // For example, you could import data from a DataTable, CSV, or another source.
            // This placeholder represents that operation.
            // --------------------------------------------------------------------
            // ImportData(workbook);

            // Refresh every pivot table in every worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (PivotTable pivot in sheet.PivotTables)
                {
                    // Update the pivot cache with the latest source data
                    pivot.RefreshData();

                    // Recalculate the pivot table values after the refresh
                    pivot.CalculateData();
                }
            }

            // Save the workbook after all pivot tables have been refreshed
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any runtime exceptions gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Placeholder for bulk data import logic (implementation depends on your data source)
    // static void ImportData(Workbook wb) { /* ... */ }
}
