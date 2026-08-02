// Title: Measure memory usage of Aspose.Cells HTML export with CSS disabled in C#
// Description: A C# console sample that loads a large Excel workbook (or creates one if missing), sets HtmlSaveOptions.DisableCss = true, records the process's private memory before and after Workbook.Save to HTML, forces garbage collection, and prints the memory delta in megabytes. Ideal for benchmarking Aspose.Cells HTML export performance.
// Keywords: Aspose.Cells memory benchmark | HTML export without CSS | C# Aspose.Cells performance | measure process memory .NET | large workbook to HTML | DisableCss Aspose.Cells | Workbook.Save memory usage | Aspose.Cells HTML conversion profiling
// Common Searches: how to measure memory usage when saving Excel to HTML with Aspose.Cells | Aspose.Cells HTML export memory consumption | disable CSS in Aspose.Cells HTML export performance | benchmark Aspose.Cells HTML conversion for large files | C# code to track memory before and after Workbook.Save
// Developer Intent: Find out how much memory is consumed by converting a large Excel workbook to HTML with CSS disabled using Aspose.Cells for .NET.
// Use Cases: Profile server memory requirements for bulk HTML exports of big workbooks. | Compare memory footprints of Aspose.Cells HTML export with and without CSS to select the optimal setting. | Validate that disabling CSS reduces memory pressure in scheduled batch conversion jobs.
// AI Prompts: Generate a C# snippet that logs private memory before and after converting an Excel file to HTML with Aspose.Cells, using HtmlSaveOptions.DisableCss and forced garbage collection. | Suggest more accurate memory‑measurement techniques for Aspose.Cells HTML export, such as using PerformanceCounter or capturing peak working set. | Create a loop that repeatedly converts the same workbook to HTML with CSS disabled and records the average memory increase.

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsMemoryMeasurement
{
    // A C# console sample that loads a large Excel workbook (or creates one if missing), sets HtmlSaveOptions.DisableCss = true, records the process's private memory before and after Workbook.Save to HTML, forces garbage collection, and prints the memory delta in megabytes. Ideal for benchmarking Aspose.Cells HTML export performance.
    class Program
    {
        static void Main()
        {
            // Path to the large workbook that will be converted to HTML.
            string sourcePath = "large.xlsx";

            try
            {
                // Ensure the source file exists; create a simple workbook if it does not.
                if (!File.Exists(sourcePath))
                {
                    Console.WriteLine($"Source file '{sourcePath}' not found. Creating a sample workbook.");
                    Workbook sample = new Workbook();
                    sample.Worksheets[0].Cells["A1"].PutValue("Sample data");
                    sample.Save(sourcePath);
                }

                // Load the workbook (uses the Workbook(string) constructor – lifecycle rule).
                using (Workbook workbook = new Workbook(sourcePath))
                {
                    // Prepare HTML save options with CSS disabled (uses HtmlSaveOptions.DisableCss property – feature rule).
                    HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                    {
                        DisableCss = true // Inline styles only, no external CSS.
                    };

                    // Measure memory usage before the conversion.
                    Process proc = Process.GetCurrentProcess();
                    long memoryBefore = proc.PrivateMemorySize64;

                    // Convert the workbook to HTML (uses Workbook.Save(string, SaveOptions) – lifecycle rule).
                    string htmlOutput = "output_no_css.html";
                    workbook.Save(htmlOutput, htmlOptions);

                    // Force garbage collection to get a more accurate post‑conversion measurement.
                    GC.Collect();
                    GC.WaitForPendingFinalizers();

                    // Measure memory usage after the conversion.
                    long memoryAfter = proc.PrivateMemorySize64;

                    // Display the memory consumption.
                    Console.WriteLine($"Memory before conversion: {memoryBefore / 1024 / 1024} MB");
                    Console.WriteLine($"Memory after  conversion: {memoryAfter / 1024 / 1024} MB");
                    Console.WriteLine($"Memory increase: {(memoryAfter - memoryBefore) / 1024 / 1024} MB");
                }
            }
            catch (Exception ex)
            {
                // Log any unexpected errors.
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
