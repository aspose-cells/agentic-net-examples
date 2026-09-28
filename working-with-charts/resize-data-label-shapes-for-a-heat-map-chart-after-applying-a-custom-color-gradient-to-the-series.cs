// Title: How to resize data label shapes in a column chart with a custom two‑color gradient using Aspose.Cells for .NET
// AI Prompts: Write C# code that creates a column chart, applies a vertical two‑color gradient to the first series, enables data labels, turns off automatic sizing, and sets each label’s width to 60 points with Aspose.Cells. | Show how to iterate over ChartPoint objects in Aspose.Cells to adjust the DataLabels.Width property after applying a gradient fill. | Explain how to recalculate the chart and save the workbook after modifying data label dimensions in Aspose.Cells for .NET.
// Common Searches: Aspose.Cells set fixed width for chart data labels after applying gradient fill | C# resize data label shape in column chart Aspose.Cells | prevent automatic label sizing in Aspose.Cells column chart | apply vertical two‑color gradient to series and adjust label size Aspose.Cells
// Tags: set data label width Aspose.Cells | apply two‑color gradient series Aspose.Cells | turn off auto sizing chart data labels Aspose.Cells | resize chart point data label shape Aspose.Cells | column chart gradient fill Aspose.Cells .NET | heat‑map style column chart Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;   // For GradientStyleType

namespace AsposeCellsHeatMapDataLabelResize
{
    // Creates a workbook, adds a column chart with sample data, applies a vertical two‑color gradient to the first series, enables data labels, turns off automatic sizing for each point, sets each label’s width to 60, recalculates the chart, and saves the file as HeatMapDataLabelResize.xlsx.
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for a heat‑map‑like chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Series1");
                sheet.Cells["C1"].PutValue("Series2");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["A4"].PutValue("C");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(30);
                sheet.Cells["B4"].PutValue(20);
                sheet.Cells["C2"].PutValue(25);
                sheet.Cells["C3"].PutValue(15);
                sheet.Cells["C4"].PutValue(35);

                // Add a column chart (HeatMap chart type is not available in Aspose.Cells)
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
                Chart chart = sheet.Charts[chartIndex];

                // Set data range for the chart
                chart.SetChartDataRange("A1:C4", true);

                // Apply a custom two‑color gradient to the first series
                Series series = chart.NSeries[0];
                series.Area.FillFormat.SetTwoColorGradient(
                    Color.LightBlue,   // start color
                    Color.DarkBlue,    // end color
                    GradientStyleType.Vertical,
                    1);                 // gradient variant

                // Enable data labels and set basic appearance
                series.DataLabels.ShowValue = true;
                series.DataLabels.Position = LabelPositionType.Center;
                // ShapeType property may not be supported in all versions; omit if unavailable
                // series.DataLabels.ShapeType = DataLabelShapeType.Rect;
                series.DataLabels.ApplyFont();

                // Resize each data label shape
                foreach (ChartPoint point in series.Points)
                {
                    // Disable auto‑fit so we can set a fixed size
                    point.DataLabels.IsResizeShapeToFitText = false;

                    // Set a custom width (height can be set similarly if needed)
                    point.DataLabels.Width = 60;
                }

                // Recalculate chart to apply changes
                chart.Calculate();

                // Save the workbook
                workbook.Save("HeatMapDataLabelResize.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
