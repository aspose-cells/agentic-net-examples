// Title: Measure and log the rendering time of a complex line chart using Aspose.Cells for .NET
// AI Prompts: Create a C# program that fills a worksheet with 1000 rows and 5 data series, adds a line chart, renders it to a 300 DPI PNG image, and prints the rendering duration in milliseconds. | Adapt the example to export the chart as a JPEG at 150 DPI and record the elapsed time in seconds instead of milliseconds. | Extend the code to generate both line and column charts, render each to an image, and compare their rendering times using Stopwatch.
// Common Searches: aspnet measure Aspose.Cells chart rendering performance | how long does ToImage take for a line chart in Aspose.Cells | benchmark chart export speed Aspose.Cells C# | log chart rendering time with Stopwatch in .NET
// Tags: chart rendering performance measurement Aspose.Cells | line chart image generation C# | elapsed time measurement .NET | high resolution chart export Aspose.Cells | performance testing complex chart rendering .NET

using System;
using System.Diagnostics;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

// Creates a workbook with 1000 rows of sinusoidal data across five series, adds a line chart, renders it to a 300‑DPI PNG image, and logs the rendering duration in milliseconds using Stopwatch.
class ChartRenderingPerformance
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate worksheet with sample data for a complex chart (e.g., 1000 rows, 5 series)
            int rows = 1000;
            int seriesCount = 5;
            for (int i = 0; i < rows; i++)
            {
                // X axis values (e.g., dates or categories)
                sheet.Cells[i, 0].PutValue(i + 1);

                // Y values for each series
                for (int s = 0; s < seriesCount; s++)
                {
                    sheet.Cells[i, s + 1].PutValue(Math.Sin((i + s) * 0.01) * 100 + s * 20);
                }
            }

            // Add a line chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Line, 5, 7, 30, 20);
            Chart chart = sheet.Charts[chartIndex];

            // Set chart title
            chart.Title.Text = "Complex Line Chart Performance Test";

            // Add series to the chart
            for (int s = 0; s < seriesCount; s++)
            {
                // Column letters start from B (index 1) for series data
                char colLetter = (char)('B' + s);
                int seriesIndex = chart.NSeries.Add($"=Sheet1!${colLetter}$1:${colLetter}${rows}", true);
                // Set X values (category axis)
                chart.NSeries[seriesIndex].XValues = $"=Sheet1!$A$1:$A${rows}";
            }

            // Prepare image options (set resolution; default format is PNG)
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                HorizontalResolution = 300,
                VerticalResolution = 300
            };

            // Measure rendering time
            Stopwatch sw = new Stopwatch();
            sw.Start();

            // Render chart directly to an image file
            chart.ToImage("RenderedChart.png", imgOptions);

            sw.Stop();

            // Log the duration
            Console.WriteLine($"Chart rendering time: {sw.ElapsedMilliseconds} ms");

            // Save the workbook (if needed)
            workbook.Save("ComplexChartWorkbook.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
