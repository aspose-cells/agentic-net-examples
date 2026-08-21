// Title: C# – Measure memory usage when exporting a large workbook to HTML with CSS using Aspose.Cells
// Description: Creates a 5,000‑row × 50‑column workbook, applies the FileCache memory setting, enables CSS in HtmlSaveOptions, records private memory before and after workbook.Save, forces garbage collection, and reports the memory delta.
// Keywords: Aspose.Cells memory profiling | HTML export memory usage | C# Aspose.Cells large workbook | HtmlSaveOptions CSS performance | FileCache memory setting | process private memory .NET | Excel to HTML conversion benchmark
// Common Searches: measure memory consumption Aspose.Cells HTML export | C# memory usage large workbook to HTML with CSS | Aspose.Cells memory profiling during Save | how to track memory before and after workbook.Save | impact of CSS on Aspose.Cells HTML conversion memory
// Developer Intent: Find out how much memory Aspose.Cells consumes when converting a massive workbook to HTML with CSS enabled.
// Use Cases: Validate that FileCache keeps the memory footprint low for huge worksheets during HTML export. | Compare memory impact of enabling vs. disabling CSS in HtmlSaveOptions for performance tuning. | Add memory checks to CI pipelines to ensure HTML conversion stays within resource limits.
// AI Prompts: Generate C# code that logs working set, private bytes, and GC collection counts before and after saving a workbook to HTML with Aspose.Cells. | Show how to capture peak memory usage using PerformanceCounter or DiagnosticSource during HTML export with CSS enabled. | Explain best practices for configuring MemorySetting and HtmlSaveOptions to minimize memory consumption in large‑scale Excel‑to‑HTML conversions.

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsMemoryMeasurement
{
    // Creates a 5,000‑row × 50‑column workbook, applies the FileCache memory setting, enables CSS in HtmlSaveOptions, records private memory before and after workbook.Save, forces garbage collection, and reports the memory delta.
    class Program
    {
        static void Main()
        {
            // Path to the generated HTML file
            string htmlPath = Path.Combine(Path.GetTempPath(), "LargeWorkbook.html");

            // Create a large workbook (e.g., 5000 rows x 50 columns)
            Workbook workbook = new Workbook();
            // Use a memory‑friendly setting for large data
            workbook.Settings.MemorySetting = MemorySetting.FileCache;

            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate the worksheet with sample data
            for (int row = 0; row < 5000; row++)
            {
                for (int col = 0; col < 50; col++)
                {
                    cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // Configure HTML save options with CSS enabled (default behavior)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                // Ensure CSS is used (inline styles disabled)
                DisableCss = false,
                // Keep CSS in separate files for clarity (optional)
                ExportWorksheetCSSSeparately = false,
                // Enable additional CSS custom properties for better performance
                EnableCssCustomProperties = true
            };

            // Measure memory before conversion
            Process proc = Process.GetCurrentProcess();
            long memoryBefore = proc.PrivateMemorySize64;

            // Convert the workbook to HTML
            workbook.Save(htmlPath, htmlOptions);

            // Force garbage collection to get a more accurate post‑conversion measurement
            GC.Collect();
            GC.WaitForPendingFinalizers();

            // Measure memory after conversion
            long memoryAfter = proc.PrivateMemorySize64;

            // Output the results
            Console.WriteLine($"HTML file saved to: {htmlPath}");
            Console.WriteLine($"Memory before conversion: {memoryBefore / 1024 / 1024} MB");
            Console.WriteLine($"Memory after conversion : {memoryAfter / 1024 / 1024} MB");
            Console.WriteLine($"Memory increase          : {(memoryAfter - memoryBefore) / 1024 / 1024} MB");

            // Clean up
            workbook.Dispose();
        }
    }
}
