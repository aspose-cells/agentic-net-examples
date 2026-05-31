using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace FunnelChartDataLabelResizeDemo
{
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the funnel chart
                sheet.Cells["A1"].PutValue("Stage");
                sheet.Cells["A2"].PutValue("Prospects");
                sheet.Cells["A3"].PutValue("Qualified");
                sheet.Cells["A4"].PutValue("Proposals");
                sheet.Cells["A5"].PutValue("Closed");

                sheet.Cells["B1"].PutValue("Count");
                sheet.Cells["B2"].PutValue(500);
                sheet.Cells["B3"].PutValue(300);
                sheet.Cells["B4"].PutValue(150);
                sheet.Cells["B5"].PutValue(80);

                // Add a funnel chart (position: rows 5‑20, columns 0‑10)
                int chartIndex = sheet.Charts.Add(ChartType.Funnel, 5, 0, 20, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data range for the chart
                chart.NSeries.Add("B2:B5", true);
                chart.NSeries.CategoryData = "A2:A5";

                // Access the first series
                Series series = chart.NSeries[0];

                // Enable data labels and set a fixed shape size
                series.DataLabels.ShowValue = true;
                // ShapeType property is optional; removed to avoid missing enum in older versions

                // Prevent automatic resizing so we can set a fixed size
                series.DataLabels.IsResizeShapeToFitText = false;

                // Set a fixed size for the data label shape (in pixels)
                series.DataLabels.WidthPixel = 80;
                series.DataLabels.HeightPixel = 30;

                // Apply the same settings to each point individually (optional)
                foreach (ChartPoint point in series.Points)
                {
                    point.DataLabels.IsResizeShapeToFitText = false;
                    point.DataLabels.WidthPixel = 80;
                    point.DataLabels.HeightPixel = 30;
                }

                // Recalculate the chart to apply layout changes
                chart.Calculate();

                // Save the workbook
                workbook.Save("FunnelChartDataLabelResize.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}