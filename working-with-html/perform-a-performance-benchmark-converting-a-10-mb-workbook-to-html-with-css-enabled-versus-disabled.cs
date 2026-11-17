// Title: C# benchmark: Aspose.Cells HTML export performance with external CSS vs inline styles
// Common Searches: how to measure Aspose.Cells HTML export time with external CSS versus inline styles in C# | C# code to compare performance of Excel to HTML conversion using HtmlSaveOptions with and without a stylesheet | benchmark Aspose.Cells HtmlSaveOptions ExportImagesAsBase64 impact on conversion speed | timing workbook.Save to HTML with CSS enabled and disabled in .NET | performance test for converting a 10 MB Excel workbook to HTML using Aspose.Cells
// Tags: Aspose.Cells HtmlSaveOptions CSS performance | HTML export external stylesheet vs inline styles | ExportImagesAsBase64 conversion speed impact | C# stopwatch benchmark Aspose.Cells | 10MB workbook HTML conversion timing

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;

// The sample loads a 10 MB Excel workbook and saves it twice to HTML using Aspose.Cells: once with the default external CSS and separate image files, and once with inline styles and images embedded as Base64. Each save operation is timed with a Stopwatch, and the elapsed milliseconds are printed, allowing a direct performance comparison between CSS‑enabled and CSS‑disabled HTML export.
class HtmlConversionBenchmark
{
    static void Main()
    {
        try
        {
            // Path to the workbook (ensure the file exists)
            string workbookPath = "input.xlsx";

            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Error: The file \"{workbookPath}\" was not found.");
                return;
            }

            // Load the workbook (common for both tests)
            Workbook workbook;
            try
            {
                workbook = new Workbook(workbookPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error loading workbook: {ex.Message}");
                return;
            }

            // -------------------------------------------------
            // Conversion with external CSS (default behavior)
            // -------------------------------------------------
            HtmlSaveOptions cssEnabledOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                // Keep images as separate files (default)
                ExportImagesAsBase64 = false
            };

            Stopwatch swCssEnabled = Stopwatch.StartNew();
            try
            {
                workbook.Save("output_with_css.html", cssEnabledOptions);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during CSS‑enabled save: {ex.Message}");
            }
            swCssEnabled.Stop();

            // -------------------------------------------------
            // Conversion with inline styles only (no external CSS)
            // -------------------------------------------------
            HtmlSaveOptions cssDisabledOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                // Embed images as Base64 to avoid external files
                ExportImagesAsBase64 = true
                // No separate CSS file is generated; styles are inlined by default
            };

            Stopwatch swCssDisabled = Stopwatch.StartNew();
            try
            {
                workbook.Save("output_without_css.html", cssDisabledOptions);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during CSS‑disabled save: {ex.Message}");
            }
            swCssDisabled.Stop();

            // Output the elapsed times
            Console.WriteLine($"Conversion with CSS enabled : {swCssEnabled.ElapsedMilliseconds} ms");
            Console.WriteLine($"Conversion with CSS disabled: {swCssDisabled.ElapsedMilliseconds} ms");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
