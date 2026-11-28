// Title: How to measure Aspose.Cells chart rendering time after disabling data label text wrapping in C#
// Common Searches: benchmark Aspose.Cells chart ToImage method in C# | measure time to export a column chart as PNG using Aspose.Cells | disable data label text wrapping to improve chart export speed Aspose.Cells | C# Stopwatch example for chart rendering performance with Aspose.Cells | compare chart rendering time with and without data label wrapping in Aspose.Cells
// Tags: benchmark chart ToImage export Aspose.Cells | disable data label text wrap Aspose.Cells chart | measure chart rendering latency C# | column chart PNG export Aspose.Cells | performance profiling Aspose.Cells chart rendering

using System;
using System.Diagnostics;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace ChartRenderingPerformance
{
    // The example creates a workbook with sample data, adds a column chart, disables data label text wrapping, measures the time required to render the chart to a PNG image using Stopwatch, and saves both the chart image and the workbook.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // -----------------------------------------------------------------
                // 1. Create a new workbook and populate it with sample data
                // -----------------------------------------------------------------
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Header
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");

                // Sample data (10 rows)
                for (int i = 2; i <= 11; i++)
                {
                    sheet.Cells[$"A{i}"].PutValue($"Item {i - 1}");
                    sheet.Cells[$"B{i}"].PutValue(i * 10); // arbitrary numbers
                }

                // -----------------------------------------------------------------
                // 2. Add a column chart to the worksheet
                // -----------------------------------------------------------------
                int chartIndex = sheet.Charts.Add(ChartType.Column, 13, 0, 30, 15);
                Chart chart = sheet.Charts[chartIndex];

                // Set data source for the chart
                chart.NSeries.Add("B2:B11", true);          // Values
                chart.NSeries.CategoryData = "A2:A11";      // Categories

                // -----------------------------------------------------------------
                // 3. Configure DataLabels and disable text wrapping
                // -----------------------------------------------------------------
                DataLabels dataLabels = chart.NSeries[0].DataLabels;
                dataLabels.ShowValue = true;
                dataLabels.IsTextWrapped = false;
                dataLabels.Text = "This is a very long data label text that would normally wrap if wrapping were enabled.";

                // -----------------------------------------------------------------
                // 4. Measure the time required to render the chart to an image
                // -----------------------------------------------------------------
                Stopwatch sw = new Stopwatch();
                sw.Start();

                string imagePath = "RenderedChart.png";

                // Render and save the chart image directly to a file (PNG format)
                chart.ToImage(imagePath, ImageType.Png);

                sw.Stop();

                Console.WriteLine($"Chart image saved to: {imagePath}");
                Console.WriteLine($"Chart rendering time (IsTextWrapped = false): {sw.ElapsedMilliseconds} ms");

                // -----------------------------------------------------------------
                // 5. Save the workbook (optional, for verification)
                // -----------------------------------------------------------------
                string workbookPath = "ChartWorkbook.xlsx";
                workbook.Save(workbookPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to: {workbookPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
