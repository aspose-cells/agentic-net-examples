// Title: Retrieve the last refresh date of an Excel PivotTable with Aspose.Cells for .NET (RefreshByWho not available)
// Common Searches: Aspose.Cells C# get pivot table last refresh date | how to read pivot table refresh metadata using Aspose.Cells .NET | C# code to obtain Excel pivot table refresh information with Aspose | Aspose.Cells retrieve who refreshed a pivot table | pivot table audit details Aspose.Cells .NET
// Tags: Aspose.Cells PivotTable RefreshDate | C# read pivot table refresh metadata | Aspose.Cells .NET pivot table audit information | Excel pivot table last refresh user limitation | Aspose.Cells missing RefreshByWho property

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot; // Provides PivotTable type

// The example loads an Excel workbook with Aspose.Cells, accesses the first worksheet's first PivotTable, and reads its RefreshDate property. It also clarifies that the RefreshByWho property (the user who performed the last refresh) is not exposed by the Aspose.Cells API.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook that contains the PivotTable
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index or name as needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet has at least one PivotTable
            if (sheet.PivotTables.Count > 0)
            {
                // Get the first PivotTable
                PivotTable pivot = sheet.PivotTables[0];

                // Retrieve the last refresh date of the PivotTable (user info not available via API)
                DateTime refreshDate = pivot.RefreshDate;

                // Output the result
                Console.WriteLine($"PivotTable last refreshed on: {refreshDate}");
            }
            else
            {
                Console.WriteLine("No PivotTable found in the worksheet.");
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
