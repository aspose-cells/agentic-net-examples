// Title: Read PivotTable RefreshDate property from an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells get pivot table refresh date from Excel file | How to read the RefreshDate of a PivotTable using Aspose.Cells .NET | Retrieve last refreshed timestamp of Excel pivot tables with Aspose.Cells in C# | Iterate worksheets to find pivot tables and obtain RefreshDate using Aspose.Cells | Aspose.Cells example for reading PivotTable RefreshDate property
// Tags: Aspose.Cells read PivotTable RefreshDate | C# load workbook and access pivot tables | Excel pivot table last refresh timestamp Aspose | Iterate worksheets for pivot tables Aspose.Cells | Handle missing input file Aspose.Cells example

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The sample loads an existing Excel file (input.xlsx) with Aspose.Cells, iterates through each worksheet to locate any PivotTable collections, and for each PivotTable reads its RefreshDate property. It outputs the worksheet name, pivot table name, and refresh date to the console while handling missing files and runtime exceptions.
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
            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets to locate pivot tables
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Get the collection of pivot tables on the current worksheet
                PivotTableCollection pivotTables = sheet.PivotTables;

                // Skip worksheets without pivot tables
                if (pivotTables == null || pivotTables.Count == 0)
                    continue;

                // Process each pivot table
                foreach (PivotTable pivot in pivotTables)
                {
                    // Retrieve the RefreshDate property
                    DateTime refreshDate = pivot.RefreshDate;

                    // Output the refresh date
                    Console.WriteLine($"Worksheet: {sheet.Name}, PivotTable: {pivot.Name}, RefreshDate: {refreshDate}");
                }
            }
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
