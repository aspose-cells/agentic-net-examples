// Title: Proportionally Resize Aspose.Cells Chart Data Label Shapes After Expanding Chart Width (.NET)
// Description: Demonstrates how to create a workbook with a column chart, capture each data label's original pixel size, double the chart's width, calculate a scaling factor, and apply the factor to the labels' WidthPixel and HeightPixel properties while disabling auto‑fit, ensuring label shapes stay proportional after the chart is resized.
// Keywords: Aspose.Cells resize chart data labels | proportional label scaling Aspose.Cells | chart width change data label size .NET | disable auto‑fit data labels Aspose.Cells | WidthPixel HeightPixel chart label | Aspose.Cells chart manipulation C# | Excel chart data label dimensions
// Common Searches: how to keep chart data label size proportional in Aspose.Cells | resize data label shapes after changing chart width .NET | Aspose.Cells set data label WidthPixel HeightPixel | scale chart labels with chart size Aspose.Cells | disable auto‑fit for chart data labels C#
// Developer Intent: Adjust the dimensions of chart data label shapes so they remain proportionally sized when the chart width is increased.
// Use Cases: Generating Excel reports where column charts are widened for better readability while preserving the visual balance of data labels. | Applying corporate design guidelines that require fixed‑pixel label dimensions regardless of chart scaling. | Automating chart layout adjustments in bulk processing scripts that modify chart sizes programmatically.
// AI Prompts: Show C# code using Aspose.Cells to proportionally resize data label shapes after expanding a chart's width, including steps to disable auto‑fit and compute a scaling factor. | Explain how to retrieve original pixel dimensions of chart data labels, calculate a width‑based scale factor, and update WidthPixel and HeightPixel values in Aspose.Cells. | Provide a step‑by‑step tutorial for maintaining consistent data label sizes when programmatically enlarging an Excel chart with Aspose.Cells for .NET.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartLabelResizeDemo
{
    // Demonstrates how to create a workbook with a column chart, capture each data label's original pixel size, double the chart's width, calculate a scaling factor, and apply the factor to the labels' WidthPixel and HeightPixel properties while disabling auto‑fit, ensuring label shapes stay proportional after the chart is resized.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["A4"].PutValue("C");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(30);

                // Add a column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
                Chart chart = sheet.Charts[chartIndex];
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Enable data labels for the first series
                Series series = chart.NSeries[0];
                series.DataLabels.ShowValue = true;
                series.DataLabels.Position = LabelPositionType.Center;

                // Calculate once to obtain original shape sizes
                chart.Calculate();

                // Store original chart width and each data label's pixel dimensions
                double originalChartWidth = chart.ChartObject.Width;
                var originalLabelSizes = new List<(int widthPx, int heightPx)>();

                foreach (ChartPoint point in series.Points)
                {
                    // Disable auto‑fit so we can control the shape size manually
                    point.DataLabels.IsResizeShapeToFitText = false;

                    // Record current pixel dimensions
                    originalLabelSizes.Add((point.DataLabels.WidthPixel, point.DataLabels.HeightPixel));
                }

                // Increase the chart width (e.g., double it)
                // Cast to int if the Width property expects an integer value
                chart.ChartObject.Width = (int)(originalChartWidth * 2.0);

                // Re‑calculate to reflect the new chart size
                chart.Calculate();

                // Determine scaling factor based on chart width change
                double scaleFactor = chart.ChartObject.Width / originalChartWidth;

                // Apply proportional resizing to each data label shape
                int idx = 0;
                foreach (ChartPoint point in series.Points)
                {
                    var (origWidthPx, origHeightPx) = originalLabelSizes[idx++];
                    point.DataLabels.WidthPixel = (int)(origWidthPx * scaleFactor);
                    point.DataLabels.HeightPixel = (int)(origHeightPx * scaleFactor);
                }

                // Save the workbook
                string outputPath = "ChartLabelResizeDemo.xlsx";

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
