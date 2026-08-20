// Title: Measure memory usage of Aspose.Cells HTML export with CSS disabled in C#
// Description: A C# console app that loads a large Excel workbook, converts it to HTML twice—once with default CSS and once with inline styles only (HtmlSaveOptions.DisableCss). It records managed memory before and after each conversion using GC.GetTotalMemory and prints the byte difference, then disposes the workbook.
// Keywords: Aspose.Cells HTML export memory | DisableCss C# | measure GC memory Aspose.Cells | large workbook to HTML performance | .NET Excel to HTML conversion | Aspose.Cells memory profiling | HTML conversion without CSS | Aspose.Cells USA | Aspose.Cells Europe
// Common Searches: how to benchmark memory usage for Aspose.Cells HTML export | Aspose.Cells DisableCss memory impact | C# measure memory when saving Excel as HTML | compare memory footprint with and without CSS in Aspose.Cells | profile Aspose.Cells HTML conversion on large workbooks
// Developer Intent: The developer wants to quantify the memory footprint of converting a large Excel file to HTML with CSS enabled versus disabled using Aspose.Cells.
// Use Cases: Evaluate whether disabling CSS reduces RAM consumption in batch HTML conversions. | Profile memory for server‑side Excel‑to‑HTML pipelines to prevent out‑of‑memory errors. | Validate that the DisableCss flag does not introduce memory leaks during repeated exports.
// AI Prompts: Write C# code that logs peak memory usage for Aspose.Cells HTML export with DisableCss true and false. | Explain how HtmlSaveOptions.DisableCss changes the rendering pipeline and memory allocation in Aspose.Cells. | Suggest a more precise technique (e.g., PerformanceCounter, dotMemory) to measure memory across multiple HTML conversions with Aspose.Cells.

using System;
using System.Diagnostics;
using Aspose.Cells;

namespace AsposeCellsMemoryMeasurement
{
    // A C# console app that loads a large Excel workbook, converts it to HTML twice—once with default CSS and once with inline styles only (HtmlSaveOptions.DisableCss). It records managed memory before and after each conversion using GC.GetTotalMemory and prints the byte difference, then disposes the workbook.
    class Program
    {
        static void Main()
        {
            // Path to the large workbook that will be converted.
            const string sourceFile = "largeWorkbook.xlsx";

            // Load the workbook (create/load rule).
            Workbook workbook = new Workbook(sourceFile);

            // Measure memory usage when CSS is enabled (default).
            long memoryBeforeCss = GC.GetTotalMemory(true);
            ConvertToHtml(workbook, "output_with_css.html", disableCss: false);
            long memoryAfterCss = GC.GetTotalMemory(true);
            Console.WriteLine($"Memory used with CSS enabled: {memoryAfterCss - memoryBeforeCss} bytes");

            // Measure memory usage when CSS is disabled (inline styles only).
            long memoryBeforeNoCss = GC.GetTotalMemory(true);
            ConvertToHtml(workbook, "output_without_css.html", disableCss: true);
            long memoryAfterNoCss = GC.GetTotalMemory(true);
            Console.WriteLine($"Memory used with CSS disabled: {memoryAfterNoCss - memoryBeforeNoCss} bytes");

            // Clean up.
            workbook.Dispose();
        }

        // Helper method that saves the workbook as HTML using the specified DisableCss setting.
        private static void ConvertToHtml(Workbook workbook, string outputPath, bool disableCss)
        {
            // Create HTML save options (create rule) and configure CSS handling.
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
            htmlOptions.DisableCss = disableCss; // Use the rule property.

            // Save the workbook as HTML (save rule).
            workbook.Save(outputPath, htmlOptions);
        }
    }
}
