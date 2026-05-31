using System;
using System.IO;
using System.Threading.Tasks;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsDemo
{
    public static class ChartAsyncHelper
    {
        // Asynchronously creates a workbook, adds sample data, creates a chart, and saves the file.
        // This method can be awaited without blocking the calling thread.
        public static async Task CreateChartAsync(string outputPath)
        {
            // Run the intensive workbook and chart creation on a background thread.
            await Task.Run(() =>
            {
                // Ensure the output directory exists.
                string? directory = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                // Create a new workbook and get the first worksheet.
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the chart.
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["A4"].PutValue("C");
                sheet.Cells["B4"].PutValue(30);

                // Add a column chart to the worksheet.
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data range for the chart.
                chart.SetChartDataRange("A1:B4", true);

                // Configure basic chart properties.
                chart.Title.Text = "Sample Column Chart";
                chart.ShowLegend = true;
                chart.Style = 2; // Built‑in style index.

                // Ensure the chart layout is calculated before saving.
                chart.Calculate();

                // Save the workbook to the specified path.
                workbook.Save(outputPath);
            });
        }
    }

    internal class Program
    {
        // Entry point for a console application.
        private static async Task Main(string[] args)
        {
            string outputFile = "AsyncChartDemo.xlsx";

            try
            {
                // Await the asynchronous chart creation.
                await ChartAsyncHelper.CreateChartAsync(outputFile);
                Console.WriteLine($"Chart workbook saved to {Path.GetFullPath(outputFile)}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error creating chart: {ex.Message}");
            }
        }
    }
}