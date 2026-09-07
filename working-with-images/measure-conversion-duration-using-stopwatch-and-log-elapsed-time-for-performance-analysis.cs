// Title: How to measure Excel to PDF conversion time with Aspose.Cells and Stopwatch in C#
// AI Prompts: Write C# code that uses System.Diagnostics.Stopwatch to time the Aspose.Cells workbook.Save call when exporting an .xlsx file to PDF and prints the elapsed milliseconds. | Show a performance benchmark for the PDF export operation of Aspose.Cells, including start/stop of Stopwatch around workbook.Save. | Create a snippet that loads an Excel workbook, converts it to PDF, and logs the conversion duration using Stopwatch.
// Common Searches: measure Aspose.Cells Excel to PDF conversion duration in C# | C# Stopwatch logging for Aspose.Cells PDF export performance | benchmark workbook.Save time when converting .xlsx to PDF with Aspose.Cells
// Tags: Aspose.Cells PDF export timing | Stopwatch workbook.Save performance | Excel to PDF conversion benchmark | measure Aspose.Cells conversion latency | C# performance logging Aspose.Cells

using System;
using System.Diagnostics;
using Aspose.Cells;

// // Loads an Excel workbook with Aspose.Cells, converts it to PDF, measures the conversion time using System.Diagnostics.Stopwatch, and writes the elapsed milliseconds to the console.
class ConversionPerformance
{
    static void Main()
    {
        // Path to the source Excel file and the output PDF file
        string sourcePath = "input.xlsx";
        string outputPath = "output.pdf";

        // Initialize a Stopwatch to measure the conversion duration
        Stopwatch stopwatch = new Stopwatch();

        // Load the workbook (Aspose.Cells create/load rule is not provided, so using standard API)
        Workbook workbook = new Workbook(sourcePath);

        // Start timing the conversion process
        stopwatch.Start();

        // Convert the workbook to PDF and save the result
        workbook.Save(outputPath, SaveFormat.Pdf);

        // Stop timing after the conversion is complete
        stopwatch.Stop();

        // Log the elapsed time
        Console.WriteLine($"Conversion completed in {stopwatch.Elapsed.TotalMilliseconds} ms.");
    }
}
