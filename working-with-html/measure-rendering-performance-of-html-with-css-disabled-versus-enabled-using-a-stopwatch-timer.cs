// Title: Benchmark Aspose.Cells HTML Export: Inline Styles vs External CSS with Stopwatch (C#)
// Description: Creates a 500×20 workbook, applies alternating bold/color styles, saves to HTML twice—once with DisableCss=true (inline styles) and once with DisableCss=false (external stylesheet)—and measures each operation with Stopwatch, outputting elapsed milliseconds.
// Keywords: Aspose.Cells HTML export performance | DisableCss benchmark | C# Stopwatch timing | inline styles vs external CSS | HTMLSaveOptions performance | Aspose.Cells rendering speed | benchmark spreadsheet to HTML
// Common Searches: Aspose.Cells measure HTML export time | How to benchmark DisableCss option in Aspose.Cells | C# Stopwatch HTMLSaveOptions performance | Compare inline style and CSS output Aspose.Cells | HTML rendering speed for large workbook Aspose
// Developer Intent: Compare the execution time of saving a workbook to HTML with inline styles only versus using an external CSS file, using Aspose.Cells and Stopwatch.
// Use Cases: Determine optimal HTML export settings for large Excel reports | Assess impact of CSS generation on server‑side rendering time | Create performance baselines for web applications that serve spreadsheet data as HTML | Guide decision to enable or disable CSS for faster page loads
// AI Prompts: Write C# code that runs the SaveHtml method multiple times and returns average elapsed time for both DisableCss settings. | Suggest how to log results to a CSV file and generate a performance chart using .NET charting libraries. | Provide a PowerShell script to invoke the compiled program and parse console output for CI reporting. | Explain how to adjust HtmlSaveOptions to further reduce HTML size while measuring speed.

using System;
using System.Diagnostics;
using Aspose.Cells;

namespace AsposeCellsHtmlPerformance
{
    // Creates a 500×20 workbook, applies alternating bold/color styles, saves to HTML twice—once with DisableCss=true (inline styles) and once with DisableCss=false (external stylesheet)—and measures each operation with Stopwatch, outputting elapsed milliseconds.
    class Program
    {
        static void Main()
        {
            // Create a workbook and populate it with sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Fill a sizable range to make the rendering measurable
            for (int row = 0; row < 500; row++)
            {
                for (int col = 0; col < 20; col++)
                {
                    Cell cell = sheet.Cells[row, col];
                    cell.PutValue($"R{row + 1}C{col + 1}");
                    // Apply some styling to increase HTML complexity
                    Style style = cell.GetStyle();
                    style.Font.IsBold = (row + col) % 2 == 0;
                    style.Font.Color = (row + col) % 3 == 0 ? System.Drawing.Color.Blue : System.Drawing.Color.Black;
                    cell.SetStyle(style);
                }
            }

            // Measure rendering with CSS disabled (inline styles only)
            TimeSpan inlineTime = SaveHtml(workbook, true, "Html_InlineStyles.html");
            Console.WriteLine($"HTML saved with DisableCss = true in {inlineTime.TotalMilliseconds} ms");

            // Measure rendering with CSS enabled (external stylesheet)
            TimeSpan cssTime = SaveHtml(workbook, false, "Html_WithCss.html");
            Console.WriteLine($"HTML saved with DisableCss = false in {cssTime.TotalMilliseconds} ms");
        }

        /// <param name="workbook">The workbook to save.</param>
        /// <param name="disableCss">If true, only inline styles are applied.</param>
        /// <param name="outputPath">The file name for the saved HTML.</param>
        /// <returns>Elapsed time for the save operation.</returns>
        private static TimeSpan SaveHtml(Workbook workbook, bool disableCss, string outputPath)
        {
            HtmlSaveOptions options = new HtmlSaveOptions
            {
                DisableCss = disableCss
            };

            Stopwatch sw = Stopwatch.StartNew();
            workbook.Save(outputPath, options);
            sw.Stop();

            return sw.Elapsed;
        }
    }
}
