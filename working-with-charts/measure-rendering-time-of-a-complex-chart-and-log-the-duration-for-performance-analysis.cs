// Title: Benchmark Chart Rendering Time with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook with 2,000 rows, add a column chart, use a Stopwatch to time the rendering process, output the elapsed milliseconds, and optionally save the file. Ideal for performance analysis of large charts in Aspose.Cells.
// Keywords: Aspose.Cells chart performance | C# chart rendering benchmark | measure chart rendering time | Aspose.Cells Stopwatch timing | render chart to MemoryStream | chart rendering speed .NET | performance testing Aspose.Cells
// Common Searches: how to benchmark Aspose.Cells chart rendering in C# | measure rendering time of large chart Aspose.Cells | log chart rendering duration with Stopwatch | Aspose.Cells performance test for column chart | render Aspose.Cells chart to memory stream and time it
// Developer Intent: The developer needs to determine how long a complex chart takes to render and record that duration for performance monitoring or optimization.
// Use Cases: Identify rendering bottlenecks when generating charts with thousands of data points. | Include rendering time metrics in automated CI/CD tests to enforce SLA thresholds. | Compare the impact of different chart types, themes, or ImageOrPrintOptions on rendering speed.
// AI Prompts: Generate C# code that measures Aspose.Cells rendering time for a pie chart with 5,000 rows and writes the result to a log file. | Show how to capture rendering durations for multiple charts in a workbook and store them in a dictionary for later analysis. | Explain how to use ImageOrPrintOptions to render a chart to a PNG byte array while timing the operation.

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

// Demonstrates how to create a workbook with 2,000 rows, add a column chart, use a Stopwatch to time the rendering process, output the elapsed milliseconds, and optionally save the file. Ideal for performance analysis of large charts in Aspose.Cells.
class ChartPerformanceDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Generate a large data set for the chart (e.g., 2000 rows)
            int dataRows = 2000;
            worksheet.Cells[0, 0].PutValue("Category");
            worksheet.Cells[0, 1].PutValue("Value");
            for (int i = 1; i <= dataRows; i++)
            {
                worksheet.Cells[i, 0].PutValue("Item " + i);
                worksheet.Cells[i, 1].PutValue(i);
            }

            // Add a column chart that uses the generated data range
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 30, 10);
            Chart chart = worksheet.Charts[chartIndex];
            chart.SetChartDataRange($"A1:B{dataRows + 1}", true);
            chart.Title.Text = "Performance Test Chart";

            // Prepare rendering options (default PNG format)
            ImageOrPrintOptions renderOptions = new ImageOrPrintOptions();

            // Start timing the rendering process
            Stopwatch timer = Stopwatch.StartNew();

            // Ensure the chart layout is up‑to‑date
            chart.Calculate();

            // Render the chart to an in‑memory stream (this triggers the actual rendering)
            using (MemoryStream imageStream = new MemoryStream())
            {
                try
                {
                    chart.ToImage(imageStream, renderOptions);
                }
                catch (Exception renderEx)
                {
                    Console.WriteLine($"Rendering error: {renderEx.Message}");
                }
            }

            // Stop timing and output the elapsed time
            timer.Stop();
            Console.WriteLine($"Chart rendering time: {timer.ElapsedMilliseconds} ms");

            // Save the workbook (optional, demonstrates standard save lifecycle)
            string outputPath = "ChartPerformance.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
