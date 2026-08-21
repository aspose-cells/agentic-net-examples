// Title: Aspose.Cells C# Benchmark: HTML Export of a 10 MB Workbook – CSS Enabled vs Disabled
// Description: Loads a ~10 MB Excel file, saves it twice as HTML using Aspose.Cells—once with external CSS (DisableCss = false) and once with inline styles only (DisableCss = true)—while timing each operation and recording the output file sizes.
// Keywords: Aspose.Cells HTML export performance | C# benchmark Excel to HTML | DisableCss Aspose.Cells | external CSS vs inline styles | large workbook HTML conversion time | HTML file size Aspose.Cells
// Common Searches: Aspose.Cells benchmark HTML export speed | HTML size difference with and without CSS in Aspose.Cells | measure Excel to HTML conversion time C# | disable CSS Aspose.Cells performance test | large Excel workbook HTML export comparison
// Developer Intent: Compare conversion speed and resulting HTML file size for a 10 MB workbook when CSS is enabled versus when it is disabled using Aspose.Cells.
// Use Cases: Select the optimal HTML export setting for high‑volume reporting pipelines. | Assess storage impact of external CSS versus inline styling in generated HTML reports. | Determine whether disabling CSS yields measurable performance gains for batch conversions.
// AI Prompts: Generate C# code that processes a directory of Excel files, converts each to HTML with both CSS enabled and disabled via Aspose.Cells, and logs time and size per file. | Create a PowerShell script to run the benchmark for multiple workbook sizes and output a CSV summary of the results. | Explain how to extend the benchmark to capture memory usage during HTML conversion with Aspose.Cells.

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;

// Loads a ~10 MB Excel file, saves it twice as HTML using Aspose.Cells—once with external CSS (DisableCss = false) and once with inline styles only (DisableCss = true)—while timing each operation and recording the output file sizes.
class HtmlConversionBenchmark
{
    static void Main()
    {
        // Path to the source workbook (approximately 10 MB)
        string sourcePath = "largeWorkbook.xlsx";

        // Load the workbook (create + load lifecycle)
        Workbook workbook = new Workbook(sourcePath);

        // -------------------- Benchmark: CSS enabled (default) --------------------
        HtmlSaveOptions optionsWithCss = new HtmlSaveOptions();
        optionsWithCss.DisableCss = false; // use external CSS

        Stopwatch sw = Stopwatch.StartNew();
        // Save the workbook as HTML with CSS
        workbook.Save("output_with_css.html", optionsWithCss);
        sw.Stop();

        long timeWithCss = sw.ElapsedMilliseconds;
        long sizeWithCss = new FileInfo("output_with_css.html").Length;

        // -------------------- Benchmark: CSS disabled (inline styles) --------------------
        HtmlSaveOptions optionsWithoutCss = new HtmlSaveOptions();
        optionsWithoutCss.DisableCss = true; // use only inline styles

        sw.Restart();
        // Save the workbook as HTML without CSS
        workbook.Save("output_without_css.html", optionsWithoutCss);
        sw.Stop();

        long timeWithoutCss = sw.ElapsedMilliseconds;
        long sizeWithoutCss = new FileInfo("output_without_css.html").Length;

        // -------------------- Results --------------------
        Console.WriteLine($"CSS enabled  : Time = {timeWithCss} ms, Size = {sizeWithCss} bytes");
        Console.WriteLine($"CSS disabled : Time = {timeWithoutCss} ms, Size = {sizeWithoutCss} bytes");
    }
}
