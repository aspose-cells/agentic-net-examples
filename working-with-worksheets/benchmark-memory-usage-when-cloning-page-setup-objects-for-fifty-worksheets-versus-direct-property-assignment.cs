// Title: Compare memory usage of cloning PageSetup via reflection versus direct property assignment for 50 worksheets in Aspose.Cells for .NET
// AI Prompts: Create a C# console application that uses Aspose.Cells to copy PageSetup settings to 50 worksheets using reflection and reports the memory consumed by the operation. | Enhance the benchmark to also record the elapsed time for both the reflection‑based copy and the manual property‑assignment approaches. | Add logic to run the memory and time measurements for different worksheet counts (e.g., 10, 50, 100) and output a summary table of the results.
// Common Searches: Aspose.Cells memory benchmark copying PageSetup properties with reflection for multiple worksheets | C# performance comparison of reflection‑based PageSetup copy versus manual assignment in Aspose.Cells | measure GC.GetTotalMemory before and after applying a template PageSetup to 50 sheets using Aspose.Cells | how to evaluate memory usage of worksheet PageSetup duplication in .NET | benchmark page setup cloning speed and memory impact in large Aspose.Cells workbooks
// Tags: reflection page setup duplication Aspose.Cells | manual page setup duplication Aspose.Cells | worksheet page setup memory analysis .NET | GC memory measurement Aspose.Cells workbook | C# page setup performance testing

using System;
using System.Reflection;
using Aspose.Cells;

// The example benchmarks two ways of applying a template PageSetup to 50 worksheets: copying all writable properties via reflection and assigning each property manually. It measures memory consumption with GC.GetTotalMemory before and after each method and prints the difference, allowing developers to compare the overhead of reflection versus direct assignment.
class PageSetupBenchmark
{
    // Copies all writable properties from one PageSetup to another using reflection
    static void CopyPageSetupProperties(PageSetup source, PageSetup target)
    {
        if (source == null || target == null) return;

        PropertyInfo[] props = typeof(PageSetup).GetProperties(BindingFlags.Public | BindingFlags.Instance);
        foreach (PropertyInfo prop in props)
        {
            if (!prop.CanWrite) continue;

            try
            {
                object value = prop.GetValue(source);
                prop.SetValue(target, value);
            }
            catch
            {
                // Ignore properties that cannot be set safely
            }
        }
    }

    static void Main()
    {
        try
        {
            // Create a workbook and configure a template worksheet's PageSetup
            Workbook wb = new Workbook();
            Worksheet templateSheet = wb.Worksheets[0];
            PageSetup templateSetup = templateSheet.PageSetup;

            templateSetup.Orientation = PageOrientationType.Landscape;
            templateSetup.PaperSize = PaperSizeType.PaperA4;
            templateSetup.FitToPagesWide = 1;
            templateSetup.FitToPagesTall = 0;
            templateSetup.PrintArea = "A1:D20";
            templateSetup.CenterHorizontally = true;
            templateSetup.CenterVertically = true;

            const int sheetCount = 50;

            // ---------- Benchmark: copying via reflection ----------
            GC.Collect();
            GC.WaitForPendingFinalizers();
            long memBeforeCopy = GC.GetTotalMemory(true);

            for (int i = 0; i < sheetCount; i++)
            {
                Worksheet ws = wb.Worksheets.Add("CopySheet" + i);
                try
                {
                    CopyPageSetupProperties(templateSetup, ws.PageSetup);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error copying properties to sheet {ws.Name}: {ex.Message}");
                }
            }

            GC.Collect();
            GC.WaitForPendingFinalizers();
            long memAfterCopy = GC.GetTotalMemory(true);
            long copyMemoryUsed = memAfterCopy - memBeforeCopy;

            // ---------- Benchmark: direct property assignment ----------
            Workbook wbDirect = new Workbook();
            GC.Collect();
            GC.WaitForPendingFinalizers();
            long memBeforeDirect = GC.GetTotalMemory(true);

            for (int i = 0; i < sheetCount; i++)
            {
                Worksheet ws = wbDirect.Worksheets.Add("DirectSheet" + i);
                PageSetup ps = ws.PageSetup;
                try
                {
                    ps.Orientation = templateSetup.Orientation;
                    ps.PaperSize = templateSetup.PaperSize;
                    ps.FitToPagesWide = templateSetup.FitToPagesWide;
                    ps.FitToPagesTall = templateSetup.FitToPagesTall;
                    ps.PrintArea = templateSetup.PrintArea;
                    ps.CenterHorizontally = templateSetup.CenterHorizontally;
                    ps.CenterVertically = templateSetup.CenterVertically;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error assigning properties to sheet {ws.Name}: {ex.Message}");
                }
            }

            GC.Collect();
            GC.WaitForPendingFinalizers();
            long memAfterDirect = GC.GetTotalMemory(true);
            long directMemoryUsed = memAfterDirect - memBeforeDirect;

            // Output results
            Console.WriteLine($"Memory used when copying via reflection for {sheetCount} sheets: {copyMemoryUsed} bytes");
            Console.WriteLine($"Memory used when assigning properties directly for {sheetCount} sheets: {directMemoryUsed} bytes");
            Console.WriteLine($"Difference (Copy - Direct): {copyMemoryUsed - directMemoryUsed} bytes");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
