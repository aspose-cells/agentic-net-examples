// Title: C# Benchmark: Measure PageSetup Update Time for 100 Worksheets with Aspose.Cells
// Description: Creates a workbook with 100 sheets, uses Stopwatch to time setting FitToPagesWide, FitToPagesTall, and Landscape orientation on each worksheet's PageSetup, outputs elapsed milliseconds, and saves the file.
// Keywords: Aspose.Cells | C# | .NET | PageSetup performance | worksheet benchmark | Stopwatch timing | bulk page layout | execution time measurement
// Common Searches: Aspose.Cells how long to set page setup for many sheets | benchmark page layout changes in .NET workbook | measure performance of PageSetup across worksheets | C# timing Aspose.Cells page setup | speed test updating worksheet page settings
// Developer Intent: The developer wants to benchmark how quickly page setup properties can be applied to a large number of worksheets.
// Use Cases: Establish a baseline before optimizing bulk page setup operations | Validate that page layout changes meet UI latency requirements in large workbooks | Compare performance of different batch update strategies or workbook compression settings
// AI Prompts: Provide a C# example that records the time taken to apply margin settings to all worksheets using Aspose.Cells. | Suggest performance‑tuning tips for bulk PageSetup modifications in Aspose.Cells. | Generate a unit test that asserts the PageSetup update for 100 sheets completes within a specified duration.

using System;
using System.Diagnostics;
using Aspose.Cells;

// Creates a workbook with 100 sheets, uses Stopwatch to time setting FitToPagesWide, FitToPagesTall, and Landscape orientation on each worksheet's PageSetup, outputs elapsed milliseconds, and saves the file.
class PageSetupPerformanceDemo
{
    static void Main()
    {
        // Create a new workbook (starts with one worksheet)
        Workbook workbook = new Workbook();

        // Add worksheets so the workbook has 100 sheets in total
        for (int i = 1; i < 100; i++)
        {
            workbook.Worksheets.Add();
        }

        // Start measuring time
        Stopwatch stopwatch = Stopwatch.StartNew();

        // Update page setup for each worksheet
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Example settings: fit to one page wide and tall, landscape orientation
            sheet.PageSetup.FitToPagesWide = 1;
            sheet.PageSetup.FitToPagesTall = 1;
            sheet.PageSetup.Orientation = PageOrientationType.Landscape;
        }

        // Stop measuring
        stopwatch.Stop();

        Console.WriteLine($"Updating page setup for 100 worksheets took {stopwatch.ElapsedMilliseconds} ms.");

        // Save the workbook (optional, demonstrates normal save workflow)
        workbook.Save("PageSetupPerformance.xlsx");
    }
}
