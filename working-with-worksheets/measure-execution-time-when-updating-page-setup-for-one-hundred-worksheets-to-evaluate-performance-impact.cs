// Title: Measure execution time of updating PageSetup properties for 100 worksheets with Aspose.Cells for .NET
// AI Prompts: Write C# code that iterates through 100 worksheets in an Aspose.Cells workbook, modifies PageSetup settings (orientation, paper size, margins, fit‑to‑page) and records the duration using Stopwatch. | Create a performance benchmark that applies bulk PageSetup changes to each sheet in a workbook and outputs the elapsed milliseconds for the operation.
// Common Searches: how long does it take to change page setup for many sheets using Aspose.Cells .NET | benchmarking worksheet page orientation updates with Aspose.Cells | measure performance of batch margin settings in a 100‑sheet workbook C# | stopwatch timing Aspose.Cells PageSetup modifications across worksheets
// Tags: Aspose.Cells bulk PageSetup update timing | C# worksheet page layout performance measurement | benchmark multiple worksheet page settings Aspose.Cells | measure page setup latency .NET | performance testing Aspose.Cells page configuration

using System;
using System.Diagnostics;
using Aspose.Cells;

// The example creates a workbook containing 100 worksheets, loops through each sheet to set orientation, paper size, margins, and fit‑to‑page options, measures the total time with Stopwatch, and prints the elapsed milliseconds, allowing developers to assess the performance impact of bulk PageSetup updates.
class PageSetupPerformanceTest
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Ensure there are 100 worksheets (the workbook initially contains one worksheet)
            for (int i = 1; i < 100; i++)
            {
                workbook.Worksheets.Add("Sheet" + (i + 1));
            }

            // Start measuring time
            Stopwatch sw = Stopwatch.StartNew();

            // Update page setup for each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                PageSetup setup = sheet.PageSetup;
                // Example modifications (adjust as needed)
                setup.Orientation = PageOrientationType.Portrait;
                setup.PaperSize = PaperSizeType.PaperA4;
                setup.LeftMargin = 0.5;   // inches
                setup.RightMargin = 0.5;  // inches
                setup.TopMargin = 0.75;   // inches
                setup.BottomMargin = 0.75; // inches
                setup.FitToPagesWide = 1;
                setup.FitToPagesTall = 0;
            }

            // Stop measuring time
            sw.Stop();

            // Output the elapsed time
            Console.WriteLine($"Time taken to update page setup for 100 worksheets: {sw.ElapsedMilliseconds} ms");

            // Optionally save the workbook (not required for timing)
            // workbook.Save("PageSetupPerformanceTest.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
