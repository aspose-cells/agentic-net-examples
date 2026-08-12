// Title: Benchmark Aspose.Cells HTML Export: CSS Enabled vs Disabled with Stopwatch (C#)
// Description: Creates a 500‑row × 20‑column workbook, applies alternating bold and colored fonts, then saves it to HTML twice—once with external CSS (DisableCss = false) and once with only inline styles (DisableCss = true). Each export is timed with System.Diagnostics.Stopwatch and the elapsed milliseconds are printed to the console.
// Keywords: Aspose.Cells | HtmlSaveOptions | DisableCss | HTML export performance | C# benchmark | Stopwatch timing | Excel to HTML | inline styles vs external CSS | .NET performance testing | document conversion speed
// Common Searches: Aspose.Cells HTML export speed test | DisableCss option performance impact | measure HTML generation time C# | benchmark Aspose.Cells HTML save | compare external CSS and inline styles Aspose.Cells
// Developer Intent: Measure and compare the time required to export a workbook to HTML with CSS enabled versus CSS disabled using Aspose.Cells.
// Use Cases: Determine the fastest HTML export setting for large worksheets in reporting dashboards. | Choose between external CSS and inline styles based on measured conversion time for automated document pipelines. | Integrate performance checks into CI/CD to catch regressions after updating Aspose.Cells.
// AI Prompts: Write C# code that runs the HTML export benchmark for multiple worksheets and returns average elapsed times for CSS enabled and disabled modes. | Provide a method that logs each export's duration and generates a console summary comparing external CSS with inline styles. | Suggest .NET libraries or techniques to visualize the benchmark results (e.g., charts, tables) after the Stopwatch measurements.

using System;
using System.Diagnostics;
using Aspose.Cells;

namespace AsposeCellsHtmlPerformance
{
    // Creates a 500‑row × 20‑column workbook, applies alternating bold and colored fonts, then saves it to HTML twice—once with external CSS (DisableCss = false) and once with only inline styles (DisableCss = true). Each export is timed with System.Diagnostics.Stopwatch and the elapsed milliseconds are printed to the console.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate the worksheet with a sizable amount of data to make rendering measurable
            for (int row = 0; row < 500; row++)
            {
                for (int col = 0; col < 20; col++)
                {
                    // Put some styled text in each cell
                    string cellName = CellsHelper.CellIndexToName(row, col);
                    worksheet.Cells[cellName].PutValue($"R{row + 1}C{col + 1}");
                    Style style = worksheet.Cells[cellName].GetStyle();
                    style.Font.IsBold = (row % 2 == 0);
                    style.Font.Color = (col % 2 == 0) ? System.Drawing.Color.Blue : System.Drawing.Color.Green;
                    worksheet.Cells[cellName].SetStyle(style);
                }
            }

            // Prepare HTML save options (common settings)
            HtmlSaveOptions options = new HtmlSaveOptions();

            // Measure rendering time with CSS enabled (default)
            options.DisableCss = false; // use external CSS
            Stopwatch swCssEnabled = Stopwatch.StartNew();
            workbook.Save("Html_With_Css.html", options);
            swCssEnabled.Stop();
            Console.WriteLine($"HTML saved with CSS enabled in {swCssEnabled.ElapsedMilliseconds} ms.");

            // Measure rendering time with CSS disabled (inline styles only)
            options.DisableCss = true; // use only inline styles
            Stopwatch swCssDisabled = Stopwatch.StartNew();
            workbook.Save("Html_With_InlineStyles.html", options);
            swCssDisabled.Stop();
            Console.WriteLine($"HTML saved with CSS disabled (inline styles) in {swCssDisabled.ElapsedMilliseconds} ms.");
        }
    }
}
