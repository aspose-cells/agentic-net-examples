// Title: Measuring execution time of bulk page setup changes for 100 worksheets with Aspose.Cells in C#
// Common Searches: c# Aspose.Cells measure time to set page orientation for multiple worksheets | benchmark page setup performance for 100 sheets using Aspose.Cells | how long does applying A4 paper size to all worksheets take in Aspose.Cells | using Stopwatch to profile bulk worksheet page setup changes in .NET | performance testing of Aspose.Cells page setup updates across many worksheets
// Tags: Aspose.Cells bulk page setup performance | C# Stopwatch Aspose.Cells profiling | measure worksheet page setup latency .NET | update orientation and paper size for multiple worksheets | benchmark Aspose.Cells page setup across worksheets

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example creates a workbook with exactly 100 worksheets, then uses System.Diagnostics.Stopwatch to time how long it takes to set each sheet's PageSetup properties (landscape orientation, A4 paper size, fit to 1 page wide, unlimited tall). After the bulk update, the elapsed milliseconds are printed, the workbook is saved as UpdatedWorkbook.xlsx, and any exceptions are caught and displayed.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Ensure the workbook contains exactly 100 worksheets
                workbook.Worksheets.Clear();
                for (int i = 0; i < 100; i++)
                {
                    workbook.Worksheets.Add("Sheet" + (i + 1));
                }

                // Start measuring execution time
                Stopwatch stopwatch = Stopwatch.StartNew();

                // Update page setup for each worksheet
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    PageSetup pageSetup = sheet.PageSetup;
                    pageSetup.Orientation = PageOrientationType.Landscape; // Correct enum
                    pageSetup.PaperSize = PaperSizeType.PaperA4;
                    pageSetup.FitToPagesWide = 1;   // Fit to 1 page wide
                    pageSetup.FitToPagesTall = 0;   // Unlimited pages tall
                }

                // Stop measuring
                stopwatch.Stop();

                Console.WriteLine($"Time taken to update page setup for 100 worksheets: {stopwatch.ElapsedMilliseconds} ms");

                // Define output file path
                string outputPath = "UpdatedWorkbook.xlsx";

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
