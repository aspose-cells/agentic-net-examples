// Title: Measure Aspose.Cells HTML export speed in C# – compare CSS enabled and disabled using Stopwatch
// AI Prompts: Write a C# console program that loads an .xlsx file with Aspose.Cells, saves it as HTML twice (once with CSS generation enabled and once disabled), and prints the elapsed milliseconds for each save using Stopwatch. | Modify the program to log the rendering times for both CSS settings into a CSV file, including the workbook name and timestamps. | Enhance the solution to iterate over all .xlsx files in a folder, exporting each to HTML with CSS on and off, and aggregate the performance data into a summary report.
// Common Searches: how to benchmark Aspose.Cells HTML export performance with CSS on and off in C# | measure time taken to save Excel as HTML using Aspose.Cells HtmlSaveOptions | disable CSS generation when converting workbook to HTML with Aspose.Cells | compare speed of HTML conversion with and without CSS using Aspose.Cells | C# Stopwatch example for Aspose.Cells HTML save timing
// Tags: Aspose.Cells HTML export with CSS disabled | C# Stopwatch timing for HtmlSaveOptions | benchmark Excel to HTML conversion Aspose.Cells | compare rendering speed CSS enabled vs disabled | measure HTML export performance Aspose.Cells

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;

// A C# console app loads an Excel workbook with Aspose.Cells, exports it to HTML twice—once with CSS generation enabled and once disabled—while measuring each export using Stopwatch and reporting the elapsed milliseconds.
class HtmlRenderPerformance
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // ---------- Rendering with CSS enabled ----------
            Stopwatch swEnabled = Stopwatch.StartNew();

            HtmlSaveOptions optionsEnabled = new HtmlSaveOptions(SaveFormat.Html)
            {
                // Ensure CSS is generated (default behavior)
                DisableCss = false
            };

            try
            {
                workbook.Save("output_css_enabled.html", optionsEnabled);
                swEnabled.Stop();
                Console.WriteLine($"HTML rendering with CSS enabled: {swEnabled.ElapsedMilliseconds} ms");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during CSS-enabled rendering: {ex.Message}");
            }

            // ---------- Rendering with CSS disabled ----------
            Stopwatch swDisabled = Stopwatch.StartNew();

            HtmlSaveOptions optionsDisabled = new HtmlSaveOptions(SaveFormat.Html)
            {
                // Disable CSS generation
                DisableCss = true
            };

            try
            {
                workbook.Save("output_css_disabled.html", optionsDisabled);
                swDisabled.Stop();
                Console.WriteLine($"HTML rendering with CSS disabled: {swDisabled.ElapsedMilliseconds} ms");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during CSS-disabled rendering: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
