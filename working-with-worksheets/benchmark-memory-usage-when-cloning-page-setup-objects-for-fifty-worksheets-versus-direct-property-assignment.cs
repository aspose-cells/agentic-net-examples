// Title: Aspose.Cells .NET – Memory Benchmark: PageSetup.Copy vs Manual Property Assignment for 50 Worksheets
// Description: A C# console program that creates a source worksheet with specific PageSetup settings, then measures and compares the memory consumption of two approaches for applying those settings to 50 worksheets: using PageSetup.Copy with a CopyOptions object versus assigning each property individually. The benchmark records memory before and after each method, saves the workbooks, and prints the memory usage and difference.
// Keywords: Aspose.Cells | .NET | C# | PageSetup.Copy | memory benchmark | worksheet cloning | property assignment | performance testing | Excel automation | GC.GetTotalMemory | workbook generation
// Common Searches: Aspose.Cells memory benchmark PageSetup.Copy | compare memory usage of PageSetup cloning vs manual settings | C# measure memory for worksheet page setup cloning | performance test Aspose.Cells page setup for many sheets | how much memory does PageSetup.Copy allocate
// Developer Intent: Identify which technique—PageSetup.Copy or explicit property assignment—consumes less memory when applied to fifty worksheets in an Aspose.Cells .NET workbook.
// Use Cases: Run the benchmark to choose the most memory‑efficient method for applying page setup across large workbooks. | Integrate the test into CI pipelines to monitor memory impact of workbook generation code. | Guide architecture decisions for memory‑constrained applications that generate Excel files with Aspose.Cells.
// AI Prompts: Explain the internal memory allocation differences between PageSetup.Copy and manual property setting in Aspose.Cells. | Suggest code enhancements to capture CPU time, run multiple iterations, and compute statistical confidence for the benchmark. | Create a similar memory‑usage benchmark for cloning worksheet styles, cells, or charts using Aspose.Cells for .NET.

using System;
using System.Diagnostics;
using Aspose.Cells;

namespace AsposeCellsMemoryBenchmark
{
    // A C# console program that creates a source worksheet with specific PageSetup settings, then measures and compares the memory consumption of two approaches for applying those settings to 50 worksheets: using PageSetup.Copy with a CopyOptions object versus assigning each property individually. The benchmark records memory before and after each method, saves the workbooks, and prints the memory usage and difference.
    class Program
    {
        static void Main()
        {
            // Prepare a source worksheet with specific page setup settings
            Workbook sourceWb = new Workbook();
            Worksheet sourceSheet = sourceWb.Worksheets[0];
            sourceSheet.PageSetup.PaperSize = PaperSizeType.PaperA4;
            sourceSheet.PageSetup.Orientation = PageOrientationType.Portrait;
            sourceSheet.PageSetup.FitToPagesWide = 1;
            sourceSheet.PageSetup.FitToPagesTall = 0;
            sourceSheet.PageSetup.PrintArea = "A1:D50";

            // ---------- Clone PageSetup using Copy ----------
            // Force garbage collection before measurement
            GC.Collect();
            GC.WaitForPendingFinalizers();
            long beforeClone = GC.GetTotalMemory(true);

            // Create a workbook and add 50 worksheets
            Workbook cloneWb = new Workbook();
            // Ensure we have at least 50 sheets (the first one already exists)
            for (int i = 1; i < 50; i++)
            {
                cloneWb.Worksheets.Add();
            }

            // Clone the page setup from the source sheet to each worksheet
            for (int i = 0; i < 50; i++)
            {
                Worksheet ws = cloneWb.Worksheets[i];
                ws.PageSetup.Copy(sourceSheet.PageSetup, new CopyOptions());
            }

            // Measure memory after cloning
            GC.Collect();
            GC.WaitForPendingFinalizers();
            long afterClone = GC.GetTotalMemory(true);
            long cloneMemoryUsed = afterClone - beforeClone;

            // Save the workbook (uses provided save rule)
            cloneWb.Save("ClonePageSetup.xlsx");

            // ---------- Direct property assignment ----------
            GC.Collect();
            GC.WaitForPendingFinalizers();
            long beforeDirect = GC.GetTotalMemory(true);

            Workbook directWb = new Workbook();
            for (int i = 1; i < 50; i++)
            {
                directWb.Worksheets.Add();
            }

            // Manually assign the same properties to each worksheet
            for (int i = 0; i < 50; i++)
            {
                PageSetup ps = directWb.Worksheets[i].PageSetup;
                ps.PaperSize = sourceSheet.PageSetup.PaperSize;
                ps.Orientation = sourceSheet.PageSetup.Orientation;
                ps.FitToPagesWide = sourceSheet.PageSetup.FitToPagesWide;
                ps.FitToPagesTall = sourceSheet.PageSetup.FitToPagesTall;
                ps.PrintArea = sourceSheet.PageSetup.PrintArea;
            }

            GC.Collect();
            GC.WaitForPendingFinalizers();
            long afterDirect = GC.GetTotalMemory(true);
            long directMemoryUsed = afterDirect - beforeDirect;

            // Save the workbook (uses provided save rule)
            directWb.Save("DirectPageSetup.xlsx");

            // Output the memory usage comparison
            Console.WriteLine("Memory used when cloning PageSetup (Copy): {0} bytes", cloneMemoryUsed);
            Console.WriteLine("Memory used when assigning properties directly: {0} bytes", directMemoryUsed);
            Console.WriteLine("Difference: {0} bytes", directMemoryUsed - cloneMemoryUsed);
        }
    }
}
