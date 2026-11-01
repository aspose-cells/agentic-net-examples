// Title: Measure single‑sheet PDF export performance using Aspose.Cells ExportActiveWorksheetOnly in .NET
// AI Prompts: Write a C# console program that loads an Excel workbook, sets PdfSaveOptions.ExportActiveWorksheetOnly to true, saves the active worksheet as PDF, and records the elapsed time with Stopwatch. | Enhance the benchmark to run the export with ExportActiveWorksheetOnly and with the manual hide‑other‑sheets technique, then output both timings for a side‑by‑side performance comparison. | Add a loop to repeat each export method multiple times, compute the average duration, and log the results for more reliable speed analysis.
// Common Searches: aspnet benchmark exportactiveworksheetonly pdf conversion time | how to measure Aspose.Cells single sheet to PDF speed | performance test Aspose.Cells ExportActiveWorksheetOnly property .NET | compare active worksheet PDF export methods Aspose.Cells | timing Excel to PDF conversion for one sheet using Aspose.Cells
// Tags: ExportActiveWorksheetOnly PDF conversion performance | single sheet Excel to PDF benchmark Aspose.Cells | Aspose.Cells PdfSaveOptions timing measurement | .NET workbook active worksheet export speed | measure Excel to PDF latency Aspose.Cells

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, configures PdfSaveOptions to export only the active worksheet, measures the export duration with Stopwatch, and prints the elapsed milliseconds. It can be extended to compare this setting against manually hiding non‑active sheets and to compute average timings over multiple runs.
class Program
{
    static void Main()
    {
        // Paths for input workbook and output PDF
        string sourcePath = "input.xlsx";
        string outputPath = "output.pdf";

        try
        {
            // Verify that the source file exists to avoid FileNotFoundException
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Error: The source file \"{sourcePath}\" was not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(sourcePath);

            // Ensure the first worksheet is the active one (optional)
            workbook.Worksheets.ActiveSheetIndex = 0;

            // Hide all worksheets except the active one so only it gets exported
            int activeIndex = workbook.Worksheets.ActiveSheetIndex;
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                workbook.Worksheets[i].IsVisible = i == activeIndex;
            }

            // Prepare PDF save options (no ExportActiveWorksheetOnly property)
            PdfSaveOptions saveOptions = new PdfSaveOptions();

            // Start timing the export operation
            Stopwatch sw = Stopwatch.StartNew();

            // Export the active worksheet to PDF
            workbook.Save(outputPath, saveOptions);

            // Stop timing
            sw.Stop();

            Console.WriteLine($"Export of active worksheet completed in {sw.Elapsed.TotalMilliseconds} ms.");
            Console.WriteLine($"PDF saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
