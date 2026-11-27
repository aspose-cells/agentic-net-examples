// Title: Measure and log the rendering time of a 3‑D column chart using Aspose.Cells for .NET
// Common Searches: how to benchmark Aspose.Cells chart ToImage method in C# | c# measure execution time of rendering a 3d column chart with Aspose.Cells | Aspose.Cells performance test for large data set chart rendering | log chart rendering duration to console using Stopwatch in .NET | render high‑resolution PNG of a complex chart and capture time with Aspose.Cells
// Tags: Aspose.Cells chart rendering benchmark | C# measure ToImage execution time | 3D column chart performance Aspose.Cells | high‑resolution PNG export Aspose.Cells | stopwatch timing chart rendering .NET

using System;
using System.Diagnostics;
using System.IO;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

namespace ChartRenderingPerformance
{
    // The example creates a workbook with 1,000 rows and five series, adds a 3‑D column chart, configures high‑resolution PNG rendering options, measures the chart's ToImage rendering time with a Stopwatch, logs the elapsed milliseconds to the console, and saves both the chart image and the workbook.
    class Program
    {
        static void Main(string[] args)
        {
            // Path where the output image will be saved
            string outputImagePath = Path.Combine(Environment.CurrentDirectory, "ComplexChart.png");

            try
            {
                // 1. Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // 2. Populate the worksheet with a large data set (e.g., 1000 rows, 5 series)
                const int rowCount = 1000;
                const int seriesCount = 5;

                // Header row
                sheet.Cells[0, 0].PutValue("Category");
                for (int s = 0; s < seriesCount; s++)
                {
                    sheet.Cells[0, s + 1].PutValue($"Series {s + 1}");
                }

                // Fill data rows
                Random rnd = new Random();
                for (int r = 1; r <= rowCount; r++)
                {
                    sheet.Cells[r, 0].PutValue($"Item {r}");
                    for (int s = 0; s < seriesCount; s++)
                    {
                        // Random values between 10 and 1000
                        sheet.Cells[r, s + 1].PutValue(rnd.Next(10, 1000));
                    }
                }

                // 3. Add a complex 3‑D column chart covering the whole data range
                int chartIndex = sheet.Charts.Add(ChartType.Column3D, 5, 0, 30, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data source for each series
                for (int s = 0; s < seriesCount; s++)
                {
                    // Data range for the series (e.g., B2:B1001, C2:C1001, ...)
                    string dataRange = CellsHelper.CellIndexToName(0, s + 1) + "2:" +
                                       CellsHelper.CellIndexToName(rowCount, s + 1);
                    chart.NSeries.Add(dataRange, true);
                }

                // Category (X) axis data
                chart.NSeries.CategoryData = "A2:A" + (rowCount + 1);

                // Optional: make the chart look a bit more complex
                chart.Title.Text = "Large Data Set 3‑D Column Chart";
                chart.ShowLegend = true;
                chart.Legend.Position = LegendPositionType.Bottom;
                chart.PlotArea.Area.ForegroundColor = Color.LightYellow;
                chart.PlotArea.Border.IsVisible = false;
                chart.ChartArea.Area.ForegroundColor = Color.LightBlue;
                chart.ChartArea.Border.IsVisible = false;
                chart.DepthPercent = 150; // deeper 3‑D effect
                chart.Elevation = 30;
                // Rotation property is not available in current API; omitted.

                // 4. Prepare image rendering options
                ImageOrPrintOptions options = new ImageOrPrintOptions
                {
                    ImageType = Aspose.Cells.Drawing.ImageType.Png,
                    HorizontalResolution = 300,
                    VerticalResolution = 300,
                    OnePagePerSheet = true // ensure the whole chart fits on one image
                };

                // 5. Measure rendering time
                Stopwatch sw = Stopwatch.StartNew();

                // Render the chart to an image file
                chart.ToImage(outputImagePath, options);

                sw.Stop();

                // 6. Log the duration
                Console.WriteLine($"Chart rendered to \"{outputImagePath}\" in {sw.Elapsed.TotalMilliseconds:F2} ms.");

                // Optional: Save the workbook for verification
                string workbookPath = Path.Combine(Environment.CurrentDirectory, "ComplexChartWorkbook.xlsx");
                workbook.Save(workbookPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to \"{workbookPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred during chart rendering: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
