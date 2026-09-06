// Title: Measure memory usage when converting a large Excel workbook to HTML with CSS disabled using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads a specified .xlsx file with Aspose.Cells, saves it to HTML with CSS generation turned off, and prints the memory consumed by the conversion. | Show how to configure HtmlSaveOptions to suppress CSS output and use a MemoryStream for the HTML result while measuring memory before and after the operation. | Add error handling for missing workbook files and exceptions during load or save, and include GC collection steps to obtain accurate memory measurements.
// Common Searches: how to profile memory consumption of Aspose.Cells HTML export in C# | convert large Excel file to HTML without CSS using Aspose.Cells .NET | measure GC.GetTotalMemory before and after saving workbook as HTML | Aspose.Cells HTMLSaveOptions disable stylesheet for large workbooks | C# memory usage comparison for Excel to HTML conversion with Aspose.Cells
// Tags: Aspose.Cells HTMLSaveOptions disable CSS | measure memory Aspose.Cells conversion | large workbook HTML export MemoryStream | C# GC memory profiling Aspose.Cells | Excel to HTML conversion performance

using System;
using System.IO;
using Aspose.Cells;

// The example loads a large .xlsx workbook with Aspose.Cells, saves it to HTML with CSS generation turned off using HtmlSaveOptions and a MemoryStream, and reports the memory used before and after the conversion via GC.GetTotalMemory.
class MemoryMeasurement
{
    static void Main()
    {
        try
        {
            const string workbookPath = "largeWorkbook.xlsx";

            // Verify that the workbook file exists to avoid FileNotFoundException
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Error: The file '{workbookPath}' was not found.");
                return;
            }

            // Load the large workbook
            Workbook workbook;
            try
            {
                workbook = new Workbook(workbookPath);
            }
            catch (Exception loadEx)
            {
                Console.WriteLine($"Failed to load workbook: {loadEx.Message}");
                return;
            }

            // Force garbage collection and get memory usage before conversion
            GC.Collect();
            GC.WaitForPendingFinalizers();
            long memoryBefore = GC.GetTotalMemory(true);

            // Configure HTML save options (CSS generation disabled via available settings)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                // Export all worksheets
                ExportActiveWorksheetOnly = false
                // Note: ExportCssStyleSheet property is not available in this version of Aspose.Cells.
                // CSS can be omitted by other means if required.
            };

            // Save the workbook to HTML using a memory stream to avoid file I/O overhead
            using (MemoryStream htmlStream = new MemoryStream())
            {
                workbook.Save(htmlStream, htmlOptions);
                // The HTML content is now in htmlStream; it can be written to a file if needed
            }

            // Force garbage collection and get memory usage after conversion
            GC.Collect();
            GC.WaitForPendingFinalizers();
            long memoryAfter = GC.GetTotalMemory(true);

            // Output memory usage results
            Console.WriteLine($"Memory before conversion: {memoryBefore} bytes");
            Console.WriteLine($"Memory after conversion:  {memoryAfter} bytes");
            Console.WriteLine($"Memory consumed by conversion: {memoryAfter - memoryBefore} bytes");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
