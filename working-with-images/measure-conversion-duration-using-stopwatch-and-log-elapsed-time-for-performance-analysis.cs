// Title: Measure Excel-to-PDF conversion time using Stopwatch with Aspose.Cells in C#
// AI Prompts: Write C# code that uses System.Diagnostics.Stopwatch to time the Workbook constructor and the Save method when converting an .xlsx file to PDF with Aspose.Cells. | Show how to output the elapsed milliseconds for loading and saving operations separately in a console application. | Add logic to calculate and display the total conversion duration by summing the load and save timings.
// Common Searches: how to profile Aspose.Cells Excel to PDF conversion speed in .NET | C# example measuring load time of Workbook with Stopwatch | log milliseconds for Aspose.Cells SaveFormat.Pdf operation | benchmark Excel workbook conversion to PDF using Aspose.Cells and Stopwatch
// Tags: Aspose.Cells workbook loading time measurement | Stopwatch timing for Aspose.Cells PDF export | performance profiling of Excel to PDF conversion .NET | measure save operation latency with Aspose.Cells | benchmark Aspose.Cells conversion duration

using System;
using System.Diagnostics;
using Aspose.Cells;

// Demonstrates using System.Diagnostics.Stopwatch to capture and log the milliseconds required to load an Excel workbook and to save it as a PDF with Aspose.Cells in a C# console app.
class Program
{
    static void Main()
    {
        // Path to the source Excel file and the output PDF file
        string sourcePath = @"C:\Data\Sample.xlsx";
        string outputPath = @"C:\Data\Sample.pdf";

        // Initialize a Stopwatch to measure conversion time
        Stopwatch sw = new Stopwatch();

        // Load the workbook (creation/loading rule is applied here)
        sw.Start(); // Start timing before loading
        Workbook workbook = new Workbook(sourcePath);
        sw.Stop(); // Stop timing after loading
        Console.WriteLine($"Workbook loaded in {sw.ElapsedMilliseconds} ms.");

        // Reset and start timing for the conversion (save) operation
        sw.Reset();
        sw.Start(); // Start timing before saving
        // Save the workbook as PDF (saving rule is applied here)
        workbook.Save(outputPath, SaveFormat.Pdf);
        sw.Stop(); // Stop timing after saving
        Console.WriteLine($"Workbook converted to PDF in {sw.ElapsedMilliseconds} ms.");

        // Total elapsed time for the whole process (optional)
        // Note: If you need the combined time, you can sum the two intervals or measure once around both operations.
    }
}
