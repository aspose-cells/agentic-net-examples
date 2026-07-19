// Title: Measure Aspose.Cells chart rendering time in C# using Stopwatch
// Description: C# sample that builds a workbook with a large data set, creates a column chart, times the ToImage rendering call with System.Diagnostics.Stopwatch, logs the elapsed milliseconds, and saves both the chart image and workbook for performance analysis.
// Keywords: Aspose.Cells chart performance | C# chart rendering benchmark | measure Aspose chart render time | Stopwatch Aspose.Cells ToImage | chart rendering latency .NET | performance testing Aspose charts
// Common Searches: how to benchmark Aspose.Cells chart rendering in C# | Aspose.Cells measure chart ToImage time | C# stopwatch chart rendering Aspose | performance test for large Aspose chart | log Aspose chart rendering duration
// Developer Intent: The developer needs to determine how long Aspose.Cells takes to render a complex chart and capture that duration for profiling or regression testing.
// Use Cases: Identify rendering bottlenecks when generating charts from massive data sets. | Compare execution times of different chart types (column, line, pie) under identical conditions. | Integrate timing metrics into CI pipelines to detect performance regressions automatically.
// AI Prompts: Generate C# code that renders an Aspose.Cells chart repeatedly with varying row counts and writes each elapsed time to a CSV file. | Explain best practices for profiling Aspose.Cells chart rendering and suggest ways to reduce latency for large workbooks. | Create a PowerShell script that runs the provided C# executable, captures the console output, and aggregates rendering times across multiple runs.

using System;
using System.Diagnostics;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsPerformanceDemo
{
    // C# sample that builds a workbook with a large data set, creates a column chart, times the ToImage rendering call with System.Diagnostics.Stopwatch, logs the elapsed milliseconds, and saves both the chart image and workbook for performance analysis.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate a large data set to make the chart complex
            int rows = 2000; // Adjust for desired complexity
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            for (int i = 0; i < rows; i++)
            {
                sheet.Cells[i + 1, 0].PutValue("Item " + (i + 1));
                sheet.Cells[i + 1, 1].PutValue(i % 100 + 1); // Sample values
            }

            // Add a column chart covering the data range
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 3, 30, 12);
            Chart chart = sheet.Charts[chartIndex];
            chart.SetChartDataRange($"A1:B{rows + 1}", true);
            chart.Title.Text = "Performance Test Chart";

            // Measure rendering time using Stopwatch
            Stopwatch sw = new Stopwatch();
            sw.Start();

            // Render the chart to an image file (this triggers the rendering process)
            chart.ToImage("RenderedChart.png");

            sw.Stop();

            // Log the duration
            Console.WriteLine($"Chart rendering time: {sw.ElapsedMilliseconds} ms");

            // Save the workbook (contains the chart and data)
            workbook.Save("PerformanceTestWorkbook.xlsx");
        }
    }
}
