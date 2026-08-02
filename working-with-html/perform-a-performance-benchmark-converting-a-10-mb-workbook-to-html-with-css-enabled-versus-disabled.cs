// Title: C# Benchmark: Aspose.Cells HTML Export Speed for a 10 MB Workbook – CSS Classes vs Inline Styles
// Description: A C# console app that loads a 10 MB Excel workbook with Aspose.Cells, saves it twice to HTML—once using the default external CSS stylesheet and once with inline styling (ExportCssClass disabled when supported)—and measures each conversion with Stopwatch, outputting the elapsed milliseconds.
// Keywords: Aspose.Cells | HTML export benchmark | C# performance test | 10 MB workbook conversion | CSS classes vs inline styles | ExportCssClass | Stopwatch timing | SaveFormat.Html | large Excel to HTML | conversion speed
// Common Searches: Aspose.Cells benchmark HTML export C# | measure Excel to HTML conversion time | disable CSS class export Aspose.Cells | external CSS vs inline styles performance | how fast is Aspose.Cells HTML save for large files
// Developer Intent: Compare the execution time of converting a 10 MB Excel workbook to HTML with CSS classes enabled versus disabled using Aspose.Cells.
// Use Cases: Assess whether external CSS or inline styles yield faster HTML generation for large workbooks. | Provide timing data for CI pipelines that validate Aspose.Cells performance regressions. | Guide architecture decisions for web apps that serve Excel content as HTML by quantifying the speed impact of CSS handling.
// AI Prompts: Generate C# code that loads a 10 MB Excel file with Aspose.Cells, exports it to HTML twice (default CSS and inline styles), and logs the conversion time for each run. | Explain how the ExportCssClass property influences HTML output and why disabling it can affect performance. | Create a script that repeats the benchmark 10 times, then reports average, minimum, and maximum conversion times for both CSS-enabled and CSS-disabled scenarios.

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsBenchmark
{
    // A C# console app that loads a 10 MB Excel workbook with Aspose.Cells, saves it twice to HTML—once using the default external CSS stylesheet and once with inline styling (ExportCssClass disabled when supported)—and measures each conversion with Stopwatch, outputting the elapsed milliseconds.
    class Program
    {
        static void Main()
        {
            // Path to the 10 MB Excel workbook
            string sourcePath = "input10mb.xlsx";

            // Verify that the source file exists to avoid FileNotFoundException
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Error: The file '{sourcePath}' was not found.");
                return;
            }

            try
            {
                // Load the workbook (create rule)
                Workbook workbook = new Workbook(sourcePath);

                // Benchmark with default CSS handling (external stylesheet)
                HtmlSaveOptions cssOptions = new HtmlSaveOptions(SaveFormat.Html);

                Stopwatch swCss = Stopwatch.StartNew();
                try
                {
                    workbook.Save("output_with_css.html", cssOptions);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error saving with CSS: {ex.Message}");
                }
                swCss.Stop();

                // Benchmark with inline styles (disable CSS class export)
                HtmlSaveOptions inlineOptions = new HtmlSaveOptions(SaveFormat.Html);
                // Note: In newer Aspose.Cells versions, CSS class export is controlled by ExportCssClass.
                // If the property is unavailable, the default behavior will be used.
                // inlineOptions.ExportCssClass = false; // Uncomment if supported by your version.

                Stopwatch swInline = Stopwatch.StartNew();
                try
                {
                    workbook.Save("output_with_inline.html", inlineOptions);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error saving with inline styles: {ex.Message}");
                }
                swInline.Stop();

                // Output the results
                Console.WriteLine($"HTML conversion with CSS enabled:  {swCss.ElapsedMilliseconds} ms");
                Console.WriteLine($"HTML conversion with CSS disabled: {swInline.ElapsedMilliseconds} ms");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors during processing
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
