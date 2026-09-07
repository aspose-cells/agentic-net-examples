// Title: Measure managed and private memory usage when converting a large .xlsx workbook to HTML with default CSS using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads a large Excel file, captures GC.GetTotalMemory and Process.PrivateMemorySize64 before and after calling Workbook.Save with HtmlSaveOptions, and outputs the memory differences in kilobytes. | Demonstrate how to invoke GC.Collect and GC.WaitForPendingFinalizers around an Aspose.Cells HTML export to obtain accurate memory measurements.
// Common Searches: C# how to benchmark memory usage of Aspose.Cells HTML conversion for big Excel files | track managed memory increase during Aspose.Cells workbook.Save to HTML | measure private memory consumption when exporting large .xlsx to HTML with Aspose.Cells | Aspose.Cells HTML export memory profiling example in .NET | record memory usage before and after converting Excel to HTML using Aspose.Cells
// Tags: Aspose.Cells HtmlSaveOptions memory profiling | C# managed memory measurement Aspose.Cells | private memory usage workbook to HTML conversion | large Excel workbook HTML export performance | default CSS styling Aspose.Cells HTML output

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;

// The sample loads a large Excel workbook, records managed and private memory before and after saving it as HTML with default CSS using Aspose.Cells, forces garbage collection to improve measurement accuracy, and prints the memory increase in kilobytes.
class Program
{
    static void Main()
    {
        // Paths for input workbook and output HTML
        string inputPath = "largeWorkbook.xlsx";
        string outputPath = "output.html";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Ensure a clean memory baseline
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();

            // Record memory usage before conversion
            long managedBefore = GC.GetTotalMemory(true);
            long privateBefore = Process.GetCurrentProcess().PrivateMemorySize64;

            // Load the workbook (large file)
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options
            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                // Export all worksheets (default is true, but set explicitly)
                ExportActiveWorksheetOnly = false
                // Note: ExportCssClass property is not available in the current Aspose.Cells version.
                // CSS styling will be applied using the default behavior.
            };

            // Perform the conversion to HTML
            workbook.Save(outputPath, saveOptions);

            // Force garbage collection again to get accurate post‑conversion memory
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();

            // Record memory usage after conversion
            long managedAfter = GC.GetTotalMemory(true);
            long privateAfter = Process.GetCurrentProcess().PrivateMemorySize64;

            // Output the memory consumption results
            Console.WriteLine($"Managed memory increase: {(managedAfter - managedBefore) / 1024} KB");
            Console.WriteLine($"Private memory increase: {(privateAfter - privateBefore) / 1024} KB");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
