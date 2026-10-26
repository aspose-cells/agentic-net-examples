// Title: Measure memory usage of Worksheet PageSetup cloning versus manual property assignment for 50 sheets with Aspose.Cells for .NET
// Common Searches: how to benchmark memory usage when copying worksheet page setup in Aspose.Cells C# | Aspose.Cells compare PageSetup.Copy and manual property setting memory impact for many sheets | C# example to profile memory consumption of adding 50 worksheets with page setup duplication | best practice for using GC.GetTotalMemory to evaluate page setup cloning performance | measure memory differences between PageSetup.Copy and individual property assignments in Aspose.Cells
// Tags: worksheet page setup copy performance | page setup property assignment impact | Aspose.Cells page setup duplication benchmark | GC.GetTotalMemory profiling Aspose.Cells worksheets | CopyOptions usage with PageSetup cloning

using System;
using Aspose.Cells;

// The sample creates a template worksheet with custom page setup settings, then benchmarks memory consumption for adding 50 worksheets by either cloning the PageSetup via the Copy method with default CopyOptions or by assigning each PageSetup property individually. It forces garbage collection before and after each operation, uses GC.GetTotalMemory to capture memory deltas, and outputs the memory differences.
class PageSetupBenchmark
{
    static void Main()
    {
        try
        {
            // Prepare a source worksheet with custom page setup settings
            Workbook sourceWb = new Workbook();
            Worksheet sourceWs = sourceWb.Worksheets[0];
            sourceWs.Name = "Template";
            sourceWs.PageSetup.Orientation = PageOrientationType.Landscape;
            sourceWs.PageSetup.PaperSize = PaperSizeType.PaperA4;
            sourceWs.PageSetup.FitToPagesWide = 1;
            sourceWs.PageSetup.FitToPagesTall = 1;
            sourceWs.PageSetup.PrintArea = "A1:D20";

            // Benchmark cloning (Copy) of PageSetup for 50 worksheets
            long memoryClone = BenchmarkClone(sourceWs);
            // Benchmark direct property assignment for 50 worksheets
            long memoryDirect = BenchmarkDirectAssign(sourceWs);

            Console.WriteLine($"Memory used after cloning PageSetup for 50 sheets: {memoryClone:N0} bytes");
            Console.WriteLine($"Memory used after direct property assignment for 50 sheets: {memoryDirect:N0} bytes");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Creates a new workbook, adds 50 worksheets and copies the PageSetup using the Copy method.
    // Returns the memory usage after the operation.
    private static long BenchmarkClone(Worksheet template)
    {
        // Force garbage collection and get baseline memory
        GC.Collect();
        GC.WaitForPendingFinalizers();
        long before = GC.GetTotalMemory(true);

        Workbook wb = new Workbook();
        // Remove the default sheet (we will add our own)
        wb.Worksheets.Clear();

        // Prepare copy options (default settings)
        CopyOptions copyOptions = new CopyOptions();

        for (int i = 0; i < 50; i++)
        {
            Worksheet ws = wb.Worksheets.Add($"Sheet_{i + 1}");
            // Clone the PageSetup from the template worksheet
            ws.PageSetup.Copy(template.PageSetup, copyOptions);
        }

        // Force garbage collection again and measure memory
        GC.Collect();
        GC.WaitForPendingFinalizers();
        long after = GC.GetTotalMemory(true);
        return after - before;
    }

    // Creates a new workbook, adds 50 worksheets and assigns each PageSetup property individually.
    // Returns the memory usage after the operation.
    private static long BenchmarkDirectAssign(Worksheet template)
    {
        // Force garbage collection and get baseline memory
        GC.Collect();
        GC.WaitForPendingFinalizers();
        long before = GC.GetTotalMemory(true);

        Workbook wb = new Workbook();
        wb.Worksheets.Clear();

        for (int i = 0; i < 50; i++)
        {
            Worksheet ws = wb.Worksheets.Add($"Sheet_{i + 1}");
            // Directly assign each property from the template
            ws.PageSetup.Orientation = template.PageSetup.Orientation;
            ws.PageSetup.PaperSize = template.PageSetup.PaperSize;
            ws.PageSetup.FitToPagesWide = template.PageSetup.FitToPagesWide;
            ws.PageSetup.FitToPagesTall = template.PageSetup.FitToPagesTall;
            ws.PageSetup.PrintArea = template.PageSetup.PrintArea;
        }

        // Force garbage collection again and measure memory
        GC.Collect();
        GC.WaitForPendingFinalizers();
        long after = GC.GetTotalMemory(true);
        return after - before;
    }
}
