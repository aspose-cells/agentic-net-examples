// Title: Measure Excel to PDF conversion time with Stopwatch in Aspose.Cells for .NET
// Common Searches: C# how to benchmark Aspose.Cells Excel to PDF conversion using Stopwatch | log execution time of workbook.Save as PDF with Aspose.Cells | measure performance of Excel to PDF export in .NET | profile Aspose.Cells PDF export latency in a C# application | record conversion duration when saving a workbook to PDF with Aspose.Cells
// Tags: Aspose.Cells workbook.Save performance timing | Stopwatch timing for PDF export | C# measurement of Aspose.Cells conversion latency | performance profiling of Aspose.Cells save operation | Excel workbook PDF export duration logging

using System;
using System.Diagnostics;
using Aspose.Cells;

// The example loads an Excel workbook, starts a Stopwatch, saves the workbook as a PDF using Aspose.Cells, stops the Stopwatch, and writes the elapsed milliseconds to the console for performance analysis.
class Program
{
    static void Main()
    {
        // Path to the source Excel file and the output PDF file
        string sourcePath = @"C:\Data\Sample.xlsx";
        string outputPath = @"C:\Data\Sample.pdf";

        // Initialize a Stopwatch to measure conversion time
        Stopwatch stopwatch = new Stopwatch();

        // Load the workbook (start timing)
        stopwatch.Start();
        Workbook workbook = new Workbook(sourcePath);

        // Perform the conversion: save the workbook as PDF
        workbook.Save(outputPath, SaveFormat.Pdf);
        // Stop timing after the save operation completes
        stopwatch.Stop();

        // Log the elapsed time
        Console.WriteLine($"Conversion completed in {stopwatch.Elapsed.TotalMilliseconds} ms.");
    }
}
