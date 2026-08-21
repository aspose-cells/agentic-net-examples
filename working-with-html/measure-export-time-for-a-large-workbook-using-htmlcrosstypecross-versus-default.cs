// Title: Benchmark Aspose.Cells HTML export: HtmlCrossType.Cross vs Default for a large workbook (C#)
// Description: Creates a 5,000‑row by 20‑column workbook, exports it to HTML twice—once with HtmlSaveOptions.HtmlCrossStringType set to Default and once to Cross—while measuring and printing the elapsed milliseconds for each run.
// Keywords: Aspose.Cells | HtmlCrossType | HTML export performance | C# benchmark | large workbook | HtmlSaveOptions | cross‑cell handling | export speed comparison | performance testing | Aspose.Cells HTML conversion
// Common Searches: Aspose.Cells HTML export benchmark C# | HtmlCrossType.Cross performance vs Default | measure HTML save time Aspose.Cells | speed test large workbook to HTML Aspose | how to profile Aspose.Cells HTML conversion
// Developer Intent: Find out which HtmlCrossStringType (Cross or Default) yields faster HTML export for a workbook with thousands of rows.
// Use Cases: Select the optimal HtmlCrossStringType for high‑volume report generation. | Assess the impact of cross‑cell handling on HTML conversion speed. | Integrate export‑time measurements into CI pipelines to catch performance regressions.
// AI Prompts: Write C# code that iterates over all HtmlCrossStringType values, logs each export duration, and summarizes the results in a table. | Generate a unit test that verifies the Cross option is not slower than Default for a 5,000‑row workbook. | Suggest code‑level optimizations to reduce HTML export time when using Aspose.Cells.

using System;
using System.Diagnostics;
using Aspose.Cells;

namespace AsposeCellsHtmlCrossPerformance
{
    // Creates a 5,000‑row by 20‑column workbook, exports it to HTML twice—once with HtmlSaveOptions.HtmlCrossStringType set to Default and once to Cross—while measuring and printing the elapsed milliseconds for each run.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (creation rule)
            Workbook workbook = new Workbook();

            // Populate the workbook with a large amount of data
            Worksheet sheet = workbook.Worksheets[0];
            const int totalRows = 5000;   // Adjust for desired size
            const int totalCols = 20;

            for (int row = 0; row < totalRows; row++)
            {
                for (int col = 0; col < totalCols; col++)
                {
                    // Insert a string that will span across cells to trigger cross‑cell handling
                    sheet.Cells[row, col].PutValue($"Row{row}_Col{col}_LongTextThatMayCrossCells");
                }
            }

            // Measure export time with HtmlCrossType.Default
            HtmlSaveOptions defaultOptions = new HtmlSaveOptions();
            defaultOptions.HtmlCrossStringType = HtmlCrossType.Default;

            Stopwatch swDefault = Stopwatch.StartNew();
            // Save workbook to HTML using the default cross type (save rule)
            workbook.Save("Export_Default.html", defaultOptions);
            swDefault.Stop();

            // Measure export time with HtmlCrossType.Cross
            HtmlSaveOptions crossOptions = new HtmlSaveOptions();
            crossOptions.HtmlCrossStringType = HtmlCrossType.Cross;

            Stopwatch swCross = Stopwatch.StartNew();
            // Save workbook to HTML using the Cross cross type (save rule)
            workbook.Save("Export_Cross.html", crossOptions);
            swCross.Stop();

            // Output the measured times
            Console.WriteLine($"Export time with HtmlCrossType.Default: {swDefault.ElapsedMilliseconds} ms");
            Console.WriteLine($"Export time with HtmlCrossType.Cross:   {swCross.ElapsedMilliseconds} ms");

            // Clean up
            workbook.Dispose();
        }
    }
}
