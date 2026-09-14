// Title: Benchmarking page‑setup updates for 100 worksheets with Aspose.Cells in C#
// AI Prompts: Write a C# console program that creates a workbook with 100 worksheets, sets each sheet's PageSetup to landscape, A4 paper size, and FitToPagesWide = 1, then measures the elapsed time with Stopwatch and prints the result. | Show how to profile the performance of bulk PageSetup modifications in Aspose.Cells by timing the operation for a hundred worksheets and outputting the duration in milliseconds. | Generate example code that applies the same page‑setup settings to every worksheet in a workbook and records the execution time using System.Diagnostics.Stopwatch.
// Common Searches: aspnet cells benchmark page setup for many worksheets | c# measure execution time updating page setup of 100 Excel sheets | performance test Aspose.Cells page setup orientation change across worksheets | how long does bulk page setup take with Aspose.Cells in .NET
// Tags: Aspose.Cells bulk page setup performance | C# Stopwatch timing worksheet settings | measure page setup latency for Excel workbook | optimize page setup updates across multiple worksheets

using System;
using System.Diagnostics;
using Aspose.Cells;

// The example creates a workbook with 100 worksheets, applies landscape orientation, A4 paper size, and FitToPagesWide = 1 to each sheet's PageSetup, measures the total time using Stopwatch, prints the elapsed milliseconds, and saves the file as PageSetupPerformanceTest.xlsx.
class PageSetupPerformanceTest
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Ensure there are 100 worksheets
        // The workbook initially contains one default worksheet
        for (int i = 1; i < 100; i++)
        {
            workbook.Worksheets.Add();
        }

        // Start timing the page setup updates
        Stopwatch sw = Stopwatch.StartNew();

        // Update page setup for each worksheet
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Example page setup changes
            sheet.PageSetup.Orientation = PageOrientationType.Landscape;
            sheet.PageSetup.PaperSize = PaperSizeType.PaperA4;
            sheet.PageSetup.FitToPagesWide = 1;
            sheet.PageSetup.FitToPagesTall = 0; // Auto
        }

        // Stop timing
        sw.Stop();

        // Output the elapsed time
        Console.WriteLine($"Time taken to update page setup for 100 worksheets: {sw.ElapsedMilliseconds} ms");

        // Save the workbook (optional, demonstrates lifecycle usage)
        workbook.Save("PageSetupPerformanceTest.xlsx");
    }
}
