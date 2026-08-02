// Title: Benchmark Aspose.Cells HTML export: HtmlCrossType.Cross vs Default (C#)
// Description: Creates a 2000‑row × 50‑column workbook filled with long strings, saves it twice—once with HtmlCrossStringType set to Default and once with Cross—while measuring each export with Stopwatch and printing the elapsed seconds.
// Keywords: Aspose.Cells | HtmlCrossType | Cross | Default | HTML export performance | benchmark | large workbook | C# | .NET | Stopwatch timing | HTML save options
// Common Searches: Aspose.Cells HtmlCrossType performance | HtmlCrossType.Cross export speed | benchmark HTML export Aspose.Cells | measure Aspose.Cells HTML save time | large workbook HTML export Aspose.Cells
// Developer Intent: Identify which HtmlCrossStringType option—Cross or Default—delivers faster HTML export for a large spreadsheet.
// Use Cases: Run a quick performance test to choose the optimal HtmlCrossStringType for server‑side report generation. | Compare export times when cross‑cell rendering is triggered by long text values. | Integrate timing logic into CI pipelines to detect regressions in HTML export speed. | Automate selection of HtmlCrossType based on workbook size and content complexity.
// AI Prompts: Write C# code that records the export duration for HtmlCrossType.Default and HtmlCrossType.Cross into a CSV file for later analysis. | Explain why HtmlCrossType.Cross can reduce HTML export time for large workbooks and suggest additional Aspose.Cells settings to improve performance. | Create an xUnit test that asserts the Cross option is not slower than Default when exporting a 2000 × 50 workbook. | Generate a PowerShell script that runs the benchmark program on multiple machines and aggregates the results.

using System;
using System.Diagnostics;
using Aspose.Cells;

namespace AsposeCellsHtmlCrossTiming
{
    // Creates a 2000‑row × 50‑column workbook filled with long strings, saves it twice—once with HtmlCrossStringType set to Default and once with Cross—while measuring each export with Stopwatch and printing the elapsed seconds.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate a large amount of data (e.g., 2000 rows x 50 columns)
            const int rows = 2000;
            const int cols = 50;
            for (int r = 0; r < rows; r++)
            {
                for (int c = 0; c < cols; c++)
                {
                    // Use a string that will cause cross‑cell rendering when it exceeds column width
                    sheet.Cells[r, c].PutValue($"Row{r}_Col{c}_LongTextToTriggerCrossCellRendering");
                }
            }

            // Measure export time with HtmlCrossType.Default
            HtmlSaveOptions defaultOptions = new HtmlSaveOptions();
            defaultOptions.HtmlCrossStringType = HtmlCrossType.Default; // Default behavior
            Stopwatch swDefault = Stopwatch.StartNew();
            workbook.Save("LargeWorkbook_Default.html", defaultOptions);
            swDefault.Stop();

            // Measure export time with HtmlCrossType.Cross
            HtmlSaveOptions crossOptions = new HtmlSaveOptions();
            crossOptions.HtmlCrossStringType = HtmlCrossType.Cross; // Optimized for large files
            Stopwatch swCross = Stopwatch.StartNew();
            workbook.Save("LargeWorkbook_Cross.html", crossOptions);
            swCross.Stop();

            // Output the timing results
            Console.WriteLine($"Export time with HtmlCrossType.Default: {swDefault.Elapsed.TotalSeconds:F2} seconds");
            Console.WriteLine($"Export time with HtmlCrossType.Cross:   {swCross.Elapsed.TotalSeconds:F2} seconds");
        }
    }
}
