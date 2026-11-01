// Title: Retrieve the first pivot table from the first worksheet using Aspose.Cells for .NET (C#)
// Common Searches: asp.net c# retrieve first pivot table from an Excel worksheet with Aspose.Cells | example code to access pivot table collection in Aspose.Cells C# | how to check pivot table count and read its name using Aspose.Cells .NET | load workbook and list all pivot tables with Aspose.Cells C# | c# Aspose.Cells get pivot tables from first worksheet
// Tags: aspocells pivot table collection handling | c# load excel workbook aspocells | first worksheet pivot table extraction aspocells | pivot table existence check aspocells c# | read pivot table name aspocells c#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot; // Required for PivotTableCollection and PivotTable

// Loads an Excel workbook, accesses the first worksheet, obtains its PivotTableCollection, checks for any pivot tables, and retrieves the first PivotTable (if present) to display its name.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";

            // Verify that the input file exists to prevent FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Get the collection of pivot tables on the worksheet
            PivotTableCollection pivotTables = worksheet.PivotTables;

            // Retrieve the first pivot table for further operations
            if (pivotTables.Count > 0)
            {
                PivotTable firstPivot = pivotTables[0];
                // Example operation: display the pivot table name
                Console.WriteLine($"First pivot table name: {firstPivot.Name}");
            }
            else
            {
                Console.WriteLine("No pivot tables found on the first worksheet.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
