// Title: Auto‑Resize Radar Chart Data Labels for Long Category Names with Aspose.Cells (C#)
// Description: Demonstrates how to create a radar chart in Aspose.Cells, show category names and values, and automatically expand data‑label shapes to fit long text by setting IsResizeShapeToFitText on the series and each chart point, then recalculating and saving the workbook.
// Keywords: Aspose.Cells radar chart | C# auto resize data labels | IsResizeShapeToFitText | long category names chart | chart label auto fit .NET | RadarChartAutoResizeDataLabels | Aspose.Cells data label formatting | dynamic label size Excel | chart.Calculate Aspose
// Common Searches: Aspose.Cells enable auto resize data label shape radar chart | C# set IsResizeShapeToFitText for radar chart series | how to fit long category names in radar chart labels | Aspose.Cells resize chart data labels automatically | radar chart label overflow solution .NET
// Developer Intent: Make radar chart data‑label shapes automatically expand to accommodate long category names.
// Use Cases: Generating a radar chart where category labels exceed default width and need dynamic resizing. | Applying IsResizeShapeToFitText to both the series and individual points for consistent label behavior. | Recalculating the chart after label adjustments to ensure correct layout before exporting the workbook.
// AI Prompts: Write C# code that creates a radar chart with long category names and enables automatic shape resizing for data labels using Aspose.Cells. | Explain the steps to apply IsResizeShapeToFitText to a radar chart series and its points in Aspose.Cells. | Show how to recalculate a radar chart after modifying data‑label settings to update the layout.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace RadarChartAutoResizeDataLabels
{
    // Demonstrates how to create a radar chart in Aspose.Cells, show category names and values, and automatically expand data‑label shapes to fit long text by setting IsResizeShapeToFitText on the series and each chart point, then recalculating and saving the workbook.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate data with long category names
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Very Long Category Name 1");
            sheet.Cells["A3"].PutValue("Extremely Long Category Name 2");
            sheet.Cells["A4"].PutValue("Super Long Category Name 3");

            sheet.Cells["B1"].PutValue("Series1");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a radar chart
            int chartIndex = sheet.Charts.Add(ChartType.Radar, 5, 0, 20, 12);
            Chart chart = sheet.Charts[chartIndex];

            // Set series data and category data
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Enable radar axis labels (category labels) – required for radar charts
            Series series = chart.NSeries[0];
            series.HasRadarAxisLabels = true;

            // Enable data labels and configure them
            series.DataLabels.ShowCategoryName = true;   // show the long category names
            series.DataLabels.ShowValue = true;         // show the values as well
            series.DataLabels.Position = LabelPositionType.Center;

            // Enable automatic resizing of the data label shape to fit the long text
            series.DataLabels.IsResizeShapeToFitText = true;

            // Optionally, ensure each individual point inherits the same setting
            foreach (ChartPoint point in series.Points)
            {
                point.DataLabels.IsResizeShapeToFitText = true;
            }

            // Recalculate the chart to apply layout changes
            chart.Calculate();

            // Save the workbook
            workbook.Save("RadarChart_AutoResizeDataLabels.xlsx");
        }
    }
}
