// Title: Log execution time of Aspose.Cells PivotTable RefreshData and CalculateData in C#
// AI Prompts: Write C# code that loads an Excel workbook with Aspose.Cells, starts a Stopwatch, calls PivotTable.RefreshData() and PivotTable.CalculateData(), stops the Stopwatch, and prints the elapsed milliseconds to the console. | Show how to output the PivotTable.RefreshDate after measuring the refresh duration and then persist the updated workbook under a different filename.
// Common Searches: C# how to measure time taken by Aspose.Cells pivot table refresh | Aspose.Cells log pivot cache refresh duration in milliseconds | benchmark RefreshData and CalculateData performance for Excel pivot tables using .NET | retrieve and display pivot table last refreshed date with Aspose.Cells | save workbook after pivot table refresh Aspose.Cells C# example
// Tags: Aspose.Cells pivot table refresh timing | C# measure RefreshData duration | log pivot cache refresh milliseconds | benchmark PivotTable.CalculateData performance | persist workbook after pivot refresh .NET

using System;
using System.Diagnostics;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotRefreshTiming
{
    // The example loads an Excel workbook, accesses the first worksheet's pivot table, uses a Stopwatch to time the RefreshData and CalculateData calls, writes the elapsed milliseconds and the pivot table's RefreshDate to the console, and then saves the modified workbook.
    class Program
    {
        static void Main()
        {
            // Path to the source workbook containing a pivot table
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Assume the pivot table is in the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            if (sheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No pivot tables found in the worksheet.");
                return;
            }

            // Get the first pivot table
            PivotTable pivotTable = sheet.PivotTables[0];

            // Measure the time taken to refresh the pivot table data
            Stopwatch sw = Stopwatch.StartNew();

            // Refresh the pivot cache and calculate the pivot table
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            sw.Stop();

            // Log the duration
            Console.WriteLine($"Pivot table refresh duration: {sw.Elapsed.TotalMilliseconds} ms");

            // Optionally display the refresh date set by the operation
            Console.WriteLine($"Pivot table last refreshed on: {pivotTable.RefreshDate}");

            // Save the workbook after refresh
            workbook.Save(outputPath);
        }
    }
}
