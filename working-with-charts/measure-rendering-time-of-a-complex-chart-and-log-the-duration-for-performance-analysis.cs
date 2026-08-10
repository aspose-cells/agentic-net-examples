// Title: Benchmark Aspose.Cells Chart Rendering Time in C#
// Description: Creates a workbook with 1,000 rows, adds a column chart, calculates its layout, and uses a Stopwatch to time the chart's rendering to a MemoryStream with ImageOrPrintOptions, then logs the elapsed milliseconds and optionally saves the file.
// Keywords: Aspose.Cells chart performance | C# chart rendering benchmark | measure rendering time Aspose | Stopwatch chart rendering | render chart to MemoryStream
// Common Searches: Aspose.Cells how to time chart rendering | C# benchmark chart generation speed | measure Aspose chart render latency | log chart rendering duration .NET
// Developer Intent: The developer needs to capture and record the time taken to render a complex chart with Aspose.Cells for performance analysis.
// Use Cases: Identify rendering bottlenecks when generating large spreadsheets. | Include rendering time metrics in automated regression tests. | Monitor chart generation latency in a live reporting service.
// AI Prompts: Write C# code that records Aspose.Cells chart rendering time for multiple charts and outputs a summary report. | Show how to compare rendering speeds of PNG and JPEG formats for Aspose.Cells charts. | Explain how to integrate chart rendering timing into a CI pipeline using Aspose.Cells.

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

// Creates a workbook with 1,000 rows, adds a column chart, calculates its layout, and uses a Stopwatch to time the chart's rendering to a MemoryStream with ImageOrPrintOptions, then logs the elapsed milliseconds and optionally saves the file.
class ChartRenderPerformance
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate worksheet with a large data set to make the chart complex
            int dataRows = 1000;
            worksheet.Cells[0, 0].PutValue("Category");
            worksheet.Cells[0, 1].PutValue("Value");
            for (int i = 1; i <= dataRows; i++)
            {
                worksheet.Cells[i, 0].PutValue($"Item {i}");
                worksheet.Cells[i, 1].PutValue(i);
            }

            // Add a chart using the ChartCollection.Add method
            int chartIdx = worksheet.Charts.Add(ChartType.Column, 5, 0, 30, 10);
            Chart chart = worksheet.Charts[chartIdx];

            // Set the data range for the chart
            chart.SetChartDataRange($"A1:B{dataRows + 1}", true);

            // Ensure chart layout is calculated before rendering
            chart.Calculate();

            // Prepare rendering options (single page). Default image format is PNG.
            ImageOrPrintOptions renderOptions = new ImageOrPrintOptions
            {
                OnePagePerSheet = true
            };

            // Measure rendering time using Stopwatch
            Stopwatch timer = Stopwatch.StartNew();

            // Render the chart to a memory stream (no file output needed for timing)
            using (MemoryStream ms = new MemoryStream())
            {
                chart.ToImage(ms, renderOptions);
            }

            timer.Stop();

            // Log the elapsed time in milliseconds
            Console.WriteLine($"Chart rendering time: {timer.ElapsedMilliseconds} ms");

            // Save the workbook (optional, demonstrates lifecycle usage)
            string outputPath = "ChartPerformance.xlsx";
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputPath)}");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
