// Title: How to proportionally resize chart data label shapes after expanding chart width using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# resize data label shape after changing chart width | scale chart data labels proportionally when increasing chart area size with Aspose.Cells | disable auto‑fit and set custom width height for chart data labels in Aspose.Cells | calculate scaling factor for data label shapes after expanding a column chart in .NET | increase column chart width while keeping data label dimensions consistent using Aspose.Cells
// Tags: chart data label shape resizing Aspose.Cells | auto‑fit disabled data labels .NET | scale chart area width with proportional label size | column chart label dimensions adjustment Aspose.Cells | calculate label scaling factor C# Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace ResizeDataLabelShapes
{
    // The example creates a workbook, adds a column chart with sample data, disables auto‑fit for data labels, records each label's original width and height, doubles the chart's width, computes a scaling factor, and then resizes each data label shape proportionally before saving the file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // -------------------------------------------------
                // 1. Populate sample data for the chart
                // -------------------------------------------------
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["A4"].PutValue("C");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(30);

                // -------------------------------------------------
                // 2. Add a column chart
                // -------------------------------------------------
                // Parameters: type, upper left row, upper left column, lower right row, lower right column
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data range for the chart
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // -------------------------------------------------
                // 3. Configure data labels for the first series
                // -------------------------------------------------
                DataLabels seriesLabels = chart.NSeries[0].DataLabels;
                seriesLabels.ShowValue = true;
                seriesLabels.Position = LabelPositionType.Center;

                // Disable auto‑fit so we can control the shape size manually
                seriesLabels.IsResizeShapeToFitText = false;

                // Define an initial width/height for the data label shape (points)
                const int initialLabelWidth = 80;
                const int initialLabelHeight = 30;

                // Store original sizes per point
                var originalLabelSizes = new System.Collections.Generic.Dictionary<int, (int width, int height)>();

                // Apply the initial size to each point's label
                for (int i = 0; i < chart.NSeries[0].Points.Count; i++)
                {
                    ChartPoint point = chart.NSeries[0].Points[i];
                    point.DataLabels.IsResizeShapeToFitText = false;
                    point.DataLabels.Width = initialLabelWidth;
                    point.DataLabels.Height = initialLabelHeight;

                    originalLabelSizes[i] = (initialLabelWidth, initialLabelHeight);
                }

                // -------------------------------------------------
                // 4. Record the original chart width (in points)
                // -------------------------------------------------
                int originalChartWidth = chart.ChartArea.Width;

                // -------------------------------------------------
                // 5. Increase the chart width (e.g., double it)
                // -------------------------------------------------
                int newChartWidth = originalChartWidth * 2;
                chart.ChartArea.Width = newChartWidth;

                // -------------------------------------------------
                // 6. Compute scaling factor and resize data label shapes proportionally
                // -------------------------------------------------
                double widthScale = (double)newChartWidth / originalChartWidth;

                for (int i = 0; i < chart.NSeries[0].Points.Count; i++)
                {
                    ChartPoint point = chart.NSeries[0].Points[i];
                    var (origW, origH) = originalLabelSizes[i];

                    point.DataLabels.Width = (int)Math.Round(origW * widthScale);
                    point.DataLabels.Height = (int)Math.Round(origH * widthScale);
                }

                // -------------------------------------------------
                // 7. Save the workbook
                // -------------------------------------------------
                string outputPath = "ResizeDataLabelsDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
