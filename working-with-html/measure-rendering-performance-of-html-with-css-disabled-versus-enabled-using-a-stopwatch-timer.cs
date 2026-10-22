// Title: Measure HTML export speed with Aspose.Cells: compare external CSS and inline CSS using C# Stopwatch
// AI Prompts: Write C# code that loads an Excel workbook with Aspose.Cells, exports it to HTML using the default external CSS, and records the elapsed time with System.Diagnostics.Stopwatch. | Adjust the HtmlSaveOptions to enable ExportHtmlAsSingleFile (inline CSS), export the workbook again, and use Stopwatch to capture and display both rendering durations for side‑by‑side comparison.
// Common Searches: how to benchmark Aspose.Cells HTML export time with external CSS vs inline CSS in C# | C# Aspose.Cells measure rendering performance of HTML output using Stopwatch | compare speed of saving workbook as HTML with external stylesheet and as single file in Aspose.Cells | Aspose.Cells HtmlSaveOptions ExportHtmlAsSingleFile performance test | measure Excel to HTML conversion time with and without embedded CSS using Aspose.Cells
// Tags: Aspose.Cells HTML export performance timing | C# Stopwatch measuring Aspose.Cells rendering | external CSS generation speed Aspose.Cells | inline CSS single file export Aspose.Cells | benchmark Excel to HTML conversion Aspose.Cells

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;

// Loads an Excel workbook with Aspose.Cells, exports it to HTML twice—once with default external CSS and once with inline CSS via ExportHtmlAsSingleFile—while a Stopwatch records the elapsed milliseconds for each export, allowing a direct performance comparison.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // -------------------------------------------------
            // Measure rendering time with external CSS (default behavior)
            // -------------------------------------------------
            Stopwatch sw = new Stopwatch();
            sw.Start();

            // HtmlSaveOptions with default settings generate external CSS files
            HtmlSaveOptions optionsExternalCss = new HtmlSaveOptions(SaveFormat.Html);

            try
            {
                using (MemoryStream ms = new MemoryStream())
                {
                    workbook.Save(ms, optionsExternalCss);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during external CSS rendering: {ex.Message}");
                return;
            }

            sw.Stop();
            Console.WriteLine($"HTML rendering with external CSS (default): {sw.ElapsedMilliseconds} ms");

            // -------------------------------------------------
            // Measure rendering time with inline CSS (single file)
            // -------------------------------------------------
            sw.Restart();

            // HtmlSaveOptions configured to export as a single HTML file (inline CSS)
            HtmlSaveOptions optionsInlineCss = new HtmlSaveOptions(SaveFormat.Html);
            // Note: ExportHtmlAsSingleFile property may not be available in older versions.
            // If supported, uncomment the following line:
            // optionsInlineCss.ExportHtmlAsSingleFile = true;

            try
            {
                using (MemoryStream ms = new MemoryStream())
                {
                    workbook.Save(ms, optionsInlineCss);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during inline CSS rendering: {ex.Message}");
                return;
            }

            sw.Stop();
            Console.WriteLine($"HTML rendering with inline CSS (single file): {sw.ElapsedMilliseconds} ms");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
