// Title: Measure memory usage of cloning PageSetup versus setting properties for 50 worksheets in Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# benchmark memory usage when copying PageSetup to many worksheets | How to compare memory impact of PageSetup.Clone versus manual property setting in .NET | Measure GC total memory for PageSetup cloning across 50 worksheets using Aspose.Cells | Performance test for setting worksheet page orientation and print area programmatically in C# | Memory consumption differences between PageSetup copy helper method and direct assignments in Aspose.Cells
// Tags: page setup cloning memory benchmark | worksheet page setup property assignment performance | Aspose.Cells memory usage comparison | GC total memory measurement for workbook | copy page setup helper method vs manual set

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // The example creates a template worksheet with specific PageSetup settings, then measures memory consumption when cloning those settings to 49 additional worksheets versus assigning the same properties individually on each of 50 worksheets, reporting the bytes used by each approach.
    class PageSetupBenchmark
    {
        static void Main()
        {
            try
            {
                // Create a workbook with a template worksheet
                Workbook wb = new Workbook();
                Worksheet templateSheet = wb.Worksheets[0];

                // Configure page setup on the template sheet
                templateSheet.PageSetup.Orientation = PageOrientationType.Landscape;
                templateSheet.PageSetup.PaperSize = PaperSizeType.PaperA4;
                templateSheet.PageSetup.FitToPagesWide = 1;
                templateSheet.PageSetup.FitToPagesTall = 0;
                templateSheet.PageSetup.PrintArea = "A1:D20";

                const int sheetCount = 50;

                // ---------- Approach 1: Clone (copy) PageSetup object ----------
                GC.Collect();
                GC.WaitForPendingFinalizers();
                long memBeforeClone = GC.GetTotalMemory(true);

                // Add worksheets and copy page setup from the template
                for (int i = 1; i < sheetCount; i++)
                {
                    // Worksheets.Add() may return an index (int) depending on the Aspose.Cells version.
                    // Retrieve the newly added worksheet via its index.
                    int newIndex = wb.Worksheets.Add();
                    Worksheet ws = wb.Worksheets[newIndex];
                    CopyPageSetup(templateSheet.PageSetup, ws.PageSetup);
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();
                long memAfterClone = GC.GetTotalMemory(true);
                long memUsedClone = memAfterClone - memBeforeClone;

                // ---------- Approach 2: Direct property assignment ----------
                Workbook wbDirect = new Workbook();

                // Ensure the workbook has the required number of worksheets
                for (int i = 1; i < sheetCount; i++)
                {
                    wbDirect.Worksheets.Add();
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();
                long memBeforeDirect = GC.GetTotalMemory(true);

                // Set page‑setup properties on each worksheet individually
                for (int i = 0; i < sheetCount; i++)
                {
                    Worksheet ws = wbDirect.Worksheets[i];
                    ws.PageSetup.Orientation = PageOrientationType.Landscape;
                    ws.PageSetup.PaperSize = PaperSizeType.PaperA4;
                    ws.PageSetup.FitToPagesWide = 1;
                    ws.PageSetup.FitToPagesTall = 0;
                    ws.PageSetup.PrintArea = "A1:D20";
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();
                long memAfterDirect = GC.GetTotalMemory(true);
                long memUsedDirect = memAfterDirect - memBeforeDirect;

                // Output results
                Console.WriteLine("Memory used with cloning (copy): {0} bytes", memUsedClone);
                Console.WriteLine("Memory used with direct assignment: {0} bytes", memUsedDirect);
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }

        // Helper method to copy relevant PageSetup properties from source to destination
        private static void CopyPageSetup(PageSetup src, PageSetup dest)
        {
            dest.Orientation = src.Orientation;
            dest.PaperSize = src.PaperSize;
            dest.FitToPagesWide = src.FitToPagesWide;
            dest.FitToPagesTall = src.FitToPagesTall;
            dest.PrintArea = src.PrintArea;
            // Extend this method if additional properties need to be copied
        }
    }
}
