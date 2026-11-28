// Title: Export an Aspose.Cells column chart to PNG while preserving data labels in C#
// Common Searches: C# Aspose.Cells save chart as PNG with data labels displayed | How to keep category and value labels on chart image using Aspose.Cells | Aspose.Cells chart ToImage preserve data labels in PNG output | Export column chart with data labels to high‑resolution PNG in .NET | Aspose.Cells ImageOrPrintOptions PNG retain data labels
// Tags: Aspose.Cells export chart PNG with data labels | ImageOrPrintOptions high‑resolution PNG rendering | preserve chart data labels during image conversion | column chart data label visibility C# | chart image generation Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing; // Needed for ImageType enum

namespace PreserveChartDataLabels
{
    // The example creates a workbook, populates it with fruit‑sales data, adds a column chart, enables both value and category data labels for each point, configures ImageOrPrintOptions for high‑resolution PNG output, and saves the chart as a PNG file while keeping the data labels visible.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the output PNG file
            const string outputPng = "ChartWithDataLabels.png";

            try
            {
                // ------------------------------------------------------------
                // 1. Create a workbook and populate it with sample data
                // ------------------------------------------------------------
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Header row
                cells["A1"].PutValue("Category");
                cells["B1"].PutValue("Value");

                // Sample data
                string[] categories = { "Apple", "Orange", "Banana", "Grape" };
                double[] values = { 1200, 850, 1500, 600 };

                for (int i = 0; i < categories.Length; i++)
                {
                    cells[$"A{2 + i}"].PutValue(categories[i]);
                    cells[$"B{2 + i}"].PutValue(values[i]);
                }

                // ------------------------------------------------------------
                // 2. Add a chart (Column) and bind it to the data range
                // ------------------------------------------------------------
                int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 15);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data source for the chart
                chart.NSeries.Add("B2:B5", true);
                chart.NSeries.CategoryData = "A2:A5";

                // Give the chart a title
                chart.Title.Text = "Fruit Sales";

                // ------------------------------------------------------------
                // 3. Enable data labels for each point in the series
                // ------------------------------------------------------------
                Series series = chart.NSeries[0];
                foreach (ChartPoint point in series.Points)
                {
                    point.DataLabels.ShowValue = true;          // Show the value
                    point.DataLabels.ShowCategoryName = true;   // Show the category name
                }

                // ------------------------------------------------------------
                // 4. Convert the chart to a PNG image while preserving data labels
                // ------------------------------------------------------------
                ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
                {
                    ImageType = ImageType.Png,          // Correct enum reference
                    HorizontalResolution = 300,
                    VerticalResolution = 300,
                    Transparent = false,
                    OnePagePerSheet = true
                };

                // Save the chart as an image
                chart.ToImage(outputPng, imgOptions);

                Console.WriteLine($"Chart image successfully saved to '{Path.GetFullPath(outputPng)}'.");
                Console.WriteLine("Data labels are visible in the generated PNG.");
            }
            catch (Exception ex)
            {
                // Report any unexpected errors – do not hide failures
                Console.Error.WriteLine("An error occurred while processing the chart:");
                Console.Error.WriteLine(ex.Message);
                Console.Error.WriteLine(ex.StackTrace);
            }
        }
    }
}
