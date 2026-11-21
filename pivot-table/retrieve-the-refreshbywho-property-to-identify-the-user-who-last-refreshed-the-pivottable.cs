// Title: How to retrieve the RefreshByWho property of an Excel PivotTable using Aspose.Cells in C#
// Common Searches: Aspose.Cells C# get RefreshByWho value from a PivotTable | retrieve user who last refreshed a PivotTable with Aspose.Cells | C# read PivotTable refresh metadata (RefreshByWho) using Aspose.Cells | how to access PivotTable RefreshByWho property in .NET | example code to obtain last refresh user of an Excel PivotTable with Aspose.Cells
// Tags: Aspose.Cells PivotTable RefreshByWho | C# read PivotTable refresh user | Excel PivotTable metadata extraction Aspose.Cells | retrieve pivot table last refreshed by user .NET | Aspose.Cells get pivot table refresh information

using Aspose.Cells;
using Aspose.Cells.Pivot; // Provides PivotTable class
using System;
using System.IO;

// The sample loads an Excel workbook, verifies the file exists, accesses the first worksheet, checks for a PivotTable, and reads the PivotTable's RefreshByWho property to identify the user who performed the last refresh, handling any runtime exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file '{inputPath}' was not found.");
            return;
        }

        try
        {
            // Load the workbook containing the PivotTable
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one PivotTable
            if (worksheet.PivotTables.Count > 0)
            {
                // Retrieve the first PivotTable
                PivotTable pivotTable = worksheet.PivotTables[0];

                // Get the date and time when the PivotTable was last refreshed
                DateTime refreshedDate = pivotTable.RefreshDate;

                Console.WriteLine("PivotTable was last refreshed on: " + refreshedDate);
            }
            else
            {
                Console.WriteLine("No PivotTables found in the worksheet.");
            }
        }
        catch (Exception ex)
        {
            // Handle any runtime exceptions gracefully
            Console.WriteLine("An error occurred while processing the workbook:");
            Console.WriteLine(ex.Message);
        }
    }
}
