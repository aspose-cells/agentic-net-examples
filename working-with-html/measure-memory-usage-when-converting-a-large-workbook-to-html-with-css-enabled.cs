// Title: How to measure managed memory usage while converting a large Excel workbook to HTML with CSS using Aspose.Cells for .NET
// AI Prompts: Show C# code that records GC.GetTotalMemory before and after saving a Workbook to HTML with CSS enabled using Aspose.Cells. | Demonstrate how to profile memory consumption of a large .xlsx to HTML conversion with HtmlSaveOptions in a .NET console application. | Provide an example that outputs the memory delta of an Aspose.Cells HTML export, including error handling and output directory preparation.
// Common Searches: C# measure memory before and after Aspose.Cells HTML export of a large workbook | how to profile heap usage when converting Excel to HTML with CSS in .NET | Aspose.Cells memory consumption example for large .xlsx to HTML conversion | track .NET managed memory during Excel to HTML conversion using HtmlSaveOptions | benchmark memory usage of Aspose.Cells HTMLSaveOptions with CSS enabled
// Tags: Aspose.Cells HTMLSaveOptions memory measurement | C# managed heap profiling Aspose.Cells conversion | large workbook to HTML performance Aspose.Cells | export Excel to HTML with CSS memory usage | GC.GetTotalMemory Aspose.Cells example

using System;
using System.IO;
using Aspose.Cells;

// The sample loads a large .xlsx file, captures the managed heap size before and after converting the workbook to HTML with CSS using HtmlSaveOptions, and prints the memory usage delta, illustrating how to monitor memory consumption during an Aspose.Cells HTML export.
class Program
{
    static void Main()
    {
        // Paths for input workbook and output HTML
        string workbookPath = @"C:\Data\LargeWorkbook.xlsx";
        string htmlOutputPath = @"C:\Data\LargeWorkbook.html";

        // Verify that the input workbook exists
        if (!File.Exists(workbookPath))
        {
            Console.WriteLine($"Error: Workbook file not found at '{workbookPath}'.");
            return;
        }

        try
        {
            // Force garbage collection to obtain a clean memory baseline
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();

            // Record memory usage before conversion (bytes)
            long memoryBefore = GC.GetTotalMemory(true);

            // Load the workbook
            Workbook workbook = new Workbook(workbookPath);

            // Configure HTML save options
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                // Export CSS is enabled by default; no explicit property needed
                ExportImagesAsBase64 = false,          // Keep images as separate files
                ExportActiveWorksheetOnly = false      // Export all worksheets
            };

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(htmlOutputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Convert the workbook to HTML
            workbook.Save(htmlOutputPath, saveOptions);

            // Record memory usage after conversion (bytes)
            long memoryAfter = GC.GetTotalMemory(true);

            // Calculate memory consumed by the conversion
            long memoryUsed = memoryAfter - memoryBefore;

            // Output the results
            Console.WriteLine($"Memory before conversion: {memoryBefore:N0} bytes");
            Console.WriteLine($"Memory after conversion:  {memoryAfter:N0} bytes");
            Console.WriteLine($"Memory used for conversion: {memoryUsed:N0} bytes");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred during conversion: {ex.Message}");
        }
    }
}
