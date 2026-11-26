// Title: How to measure memory usage while converting a large Excel workbook to HTML with embedded CSS using Aspose.Cells for .NET
// Common Searches: C# code to profile memory consumption of Aspose.Cells HTML export for a 5000x50 worksheet | measure private memory before and after saving workbook as HTML with CSS using Aspose.Cells | benchmark Aspose.Cells HTMLSaveOptions performance on large workbooks in .NET | track GC heap growth during Excel to HTML conversion with embedded CSS in C# | log memory increase when exporting Excel to HTML with Aspose.Cells
// Tags: Aspose.Cells HTML export memory profiling | large Excel workbook to HTML conversion | embedded CSS in Aspose.Cells HTML output | C# process private memory measurement | benchmark Aspose.Cells HTMLSaveOptions

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;

// The example creates a 5,000‑row by 50‑column workbook, fills each cell with deterministic data, configures HtmlSaveOptions to embed CSS, and then records the process's private memory and GC heap size before and after saving the workbook as an HTML file. It reports conversion time and the memory delta, helping developers benchmark and profile memory usage for large Excel‑to‑HTML conversions with Aspose.Cells.
class Program
{
    static void Main()
    {
        // Path for the generated HTML file
        string outputHtmlPath = Path.Combine(Directory.GetCurrentDirectory(), "LargeWorkbook.html");

        // Create a large workbook (e.g., 5000 rows x 50 columns)
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        const int totalRows = 5000;
        const int totalCols = 50;

        Console.WriteLine($"Generating workbook with {totalRows} rows and {totalCols} columns...");

        // Fill the worksheet with sample data
        for (int row = 0; row < totalRows; row++)
        {
            for (int col = 0; col < totalCols; col++)
            {
                // Simple deterministic data to avoid randomness overhead
                sheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
            }
        }

        // Configure HTML save options with CSS enabled (embed CSS into the HTML)
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions
        {
            ExportWorksheetCSSSeparately = false, // embed CSS
            ExportGridLines = true,
            ExportImagesAsBase64 = true,
            ExportActiveWorksheetOnly = true
        };

        // Warm up: force a garbage collection before measurement
        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();

        // Record memory usage before conversion
        Process currentProcess = Process.GetCurrentProcess();
        long memoryBefore = currentProcess.PrivateMemorySize64;
        long gcMemoryBefore = GC.GetTotalMemory(forceFullCollection: true);

        Console.WriteLine($"Memory before conversion: {memoryBefore / 1024 / 1024} MB (Private), {gcMemoryBefore / 1024 / 1024} MB (GC)");

        // Measure conversion time
        Stopwatch sw = Stopwatch.StartNew();

        try
        {
            // Save workbook as HTML
            workbook.Save(outputHtmlPath, htmlOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error during HTML conversion: " + ex.Message);
            return;
        }

        sw.Stop();

        // Record memory usage after conversion
        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();

        long memoryAfter = currentProcess.PrivateMemorySize64;
        long gcMemoryAfter = GC.GetTotalMemory(forceFullCollection: true);

        Console.WriteLine($"HTML saved to: {outputHtmlPath}");
        Console.WriteLine($"Conversion time: {sw.Elapsed.TotalSeconds:F2} seconds");
        Console.WriteLine($"Memory after conversion: {memoryAfter / 1024 / 1024} MB (Private), {gcMemoryAfter / 1024 / 1024} MB (GC)");
        Console.WriteLine($"Memory increase (Private): {(memoryAfter - memoryBefore) / 1024 / 1024} MB");
        Console.WriteLine($"Memory increase (GC): {(gcMemoryAfter - gcMemoryBefore) / 1024 / 1024} MB");
    }
}
