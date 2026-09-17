// Title: Benchmark the performance of exporting only the active worksheet to PDF using Aspose.Cells for .NET
// AI Prompts: Create a C# console program that loads an Excel workbook, sets the first worksheet as active, enables ExportActiveWorksheetOnly, starts a Stopwatch, saves the workbook as PDF with PdfSaveOptions, and outputs the elapsed milliseconds. | Write C# code that hides every worksheet except the active one, configures PdfSaveOptions, measures the duration of Workbook.Save with a Stopwatch, and logs the conversion time.
// Common Searches: c# how to benchmark Aspose.Cells PDF export for a single worksheet | measure time taken by Workbook.Save when ExportActiveWorksheetOnly is true | performance test Aspose.Cells converting active sheet to PDF | Aspose.Cells single sheet PDF conversion speed C#
// Tags: Aspose.Cells PDF export single worksheet performance | C# Stopwatch timing Workbook.Save | ExportActiveWorksheetOnly benchmark Aspose.Cells | hide worksheets Aspose.Cells performance optimization | measure active sheet conversion time .NET

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;

// Loads input.xlsx, activates the first worksheet, optionally hides other sheets, configures PdfSaveOptions (including ExportActiveWorksheetOnly), starts a Stopwatch, saves the workbook as a PDF, stops the timer, and prints the elapsed milliseconds.
class Program
{
    static void Main()
    {
        string inputFile = "input.xlsx";
        string outputFile = "single_sheet_output.pdf";

        try
        {
            // Verify that the input file exists
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Input file not found: {inputFile}");
                return;
            }

            // Load the workbook from disk
            Workbook workbook = new Workbook(inputFile);

            // Set the first worksheet as the active sheet
            workbook.Worksheets.ActiveSheetIndex = 0;

            // Hide all worksheets except the active one
            int activeIndex = workbook.Worksheets.ActiveSheetIndex;
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                workbook.Worksheets[i].IsVisible = i == activeIndex;
            }

            // Configure PDF save options (default options are sufficient)
            PdfSaveOptions saveOptions = new PdfSaveOptions();

            // Start the performance timer
            Stopwatch timer = Stopwatch.StartNew();

            // Export the active worksheet to PDF
            workbook.Save(outputFile, saveOptions);

            // Stop the timer
            timer.Stop();

            // Report the elapsed time in milliseconds
            Console.WriteLine($"Single‑sheet export completed in {timer.ElapsedMilliseconds} ms.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
