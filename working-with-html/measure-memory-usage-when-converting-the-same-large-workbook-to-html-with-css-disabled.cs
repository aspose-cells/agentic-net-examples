// Title: Measure memory usage of Aspose.Cells HTML export with CSS disabled in C#
// Common Searches: C# how to profile memory consumption when converting a large Excel workbook to HTML with Aspose.Cells | Aspose.Cells HTML export without CSS memory usage measurement | Detecting memory leaks during repeated Aspose.Cells HTML saves in .NET | Using GC.GetTotalMemory to benchmark Aspose.Cells HTML conversion performance | Retrieve process private memory size after Aspose.Cells HTML conversion in C#
// Tags: Aspose.Cells HTML export memory profiling | disable CSS in Aspose.Cells HtmlSaveOptions | GC.GetTotalMemory Aspose.Cells conversion | memory leak detection Aspose.Cells HTML save | process private memory Aspose.Cells .NET

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;

// The example loads a large Excel workbook with Aspose.Cells, saves it twice to HTML using default HtmlSaveOptions (which omit external CSS), and measures managed memory before and after each save with GC.GetTotalMemory. It forces garbage collection between runs, reports the memory delta for each conversion, and finally displays the process's private memory size, handling missing file and other runtime errors.
class MemoryUsageHtmlConversion
{
    static void Main()
    {
        // Path to the large workbook file
        string workbookPath = @"C:\Path\To\LargeWorkbook.xlsx";

        // Verify that the workbook file exists to avoid FileNotFoundException
        if (!File.Exists(workbookPath))
        {
            Console.WriteLine($"Error: The workbook file was not found at '{workbookPath}'.");
            return;
        }

        Workbook workbook = null;
        try
        {
            // Load the workbook (using the standard Aspose.Cells load method)
            workbook = new Workbook(workbookPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading workbook: {ex.Message}");
            return;
        }

        // Configure HTML save options (default options are used here)
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

        // First conversion: measure memory before and after
        try
        {
            long memoryBeforeFirst = GC.GetTotalMemory(true);
            workbook.Save("LargeWorkbook_NoCss_1.html", htmlOptions);
            GC.Collect();
            long memoryAfterFirst = GC.GetTotalMemory(true);
            Console.WriteLine($"First conversion memory usage: {memoryAfterFirst - memoryBeforeFirst} bytes");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error during first HTML conversion: {ex.Message}");
        }

        // Second conversion: measure memory again to see if there is a leak or increase
        try
        {
            long memoryBeforeSecond = GC.GetTotalMemory(true);
            workbook.Save("LargeWorkbook_NoCss_2.html", htmlOptions);
            GC.Collect();
            long memoryAfterSecond = GC.GetTotalMemory(true);
            Console.WriteLine($"Second conversion memory usage: {memoryAfterSecond - memoryBeforeSecond} bytes");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error during second HTML conversion: {ex.Message}");
        }

        // Optional: display total process private memory for reference
        try
        {
            Process currentProcess = Process.GetCurrentProcess();
            Console.WriteLine($"Process private memory size: {currentProcess.PrivateMemorySize64} bytes");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error retrieving process memory information: {ex.Message}");
        }
    }
}
