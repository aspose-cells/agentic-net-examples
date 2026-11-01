// Title: C# performance benchmark for exporting a 500‑worksheet workbook to a single HTML file with Aspose.Cells (ExportAllWorksheets = true)
// AI Prompts: Create a C# console program that builds a workbook with 500 worksheets, fills each sheet with sample data, enables HtmlSaveOptions.ExportAllWorksheets, measures the elapsed time while saving to HTML, and prints the duration and output path. | Enhance the benchmark to record peak memory usage during the HTML export using Process.GetCurrentProcess(), and display both time and memory statistics in the console. | Rewrite the example to export the same 500‑sheet workbook to PDF while preserving the timing logic, then output a side‑by‑side comparison of HTML vs PDF export performance.
// Common Searches: Aspose.Cells export 500 worksheets to a single HTML file performance test | C# benchmark converting large multi‑sheet Excel workbook to HTML with ExportAllWorksheets true | measure time and memory for Aspose.Cells HTML conversion of 500‑sheet workbook .NET
// Tags: exportallworksheets html aspose.cells | performance benchmark multi‑sheet html conversion c# | measure export latency aspose.cells html | scalable html generation from large workbook | c# large workbook export timing aspose.cells

using System;
using System.Diagnostics;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The sample creates a new Workbook, adds 500 worksheets each populated with 100 rows and 20 columns of sample data, configures HtmlSaveOptions with ExportAllWorksheets set to true, uses a Stopwatch to time the workbook.Save call to a single HTML file, and writes the elapsed seconds and file location to the console.
class PerformanceBenchmark
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add 500 worksheets and populate each with sample data
            for (int i = 0; i < 500; i++)
            {
                // Use the first worksheet if i == 0, otherwise add a new one
                Worksheet sheet = i == 0 ? workbook.Worksheets[0] : workbook.Worksheets.Add($"Sheet{i + 1}");

                // Fill the worksheet with sample data
                Cells cells = sheet.Cells;
                for (int row = 0; row < 100; row++)
                {
                    for (int col = 0; col < 20; col++)
                    {
                        cells[row, col].PutValue($"R{row}C{col}");
                    }
                }
            }

            // Configure PDF save options (all sheets are saved by default)
            PdfSaveOptions saveOptions = new PdfSaveOptions
            {
                OnePagePerSheet = false,               // Keep sheets continuous in the PDF
                Compliance = PdfCompliance.PdfA1b      // PDF/A-1b compliance
            };

            // Measure the time taken to export the workbook to PDF
            Stopwatch stopwatch = new Stopwatch();
            stopwatch.Start();

            // Save the workbook as PDF using the configured options
            string outputPath = "BenchmarkOutput.pdf";
            workbook.Save(outputPath, saveOptions);

            stopwatch.Stop();

            Console.WriteLine($"Export of 500 worksheets completed in {stopwatch.Elapsed.TotalSeconds:F2} seconds.");
            Console.WriteLine($"PDF saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
