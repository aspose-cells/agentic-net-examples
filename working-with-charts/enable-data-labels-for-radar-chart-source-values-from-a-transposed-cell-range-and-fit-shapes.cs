using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsRadarChartDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            // Column A – categories (radar axes)
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Cat1");
            sheet.Cells["A3"].PutValue("Cat2");
            sheet.Cells["A4"].PutValue("Cat3");

            // Column B – series values
            sheet.Cells["B1"].PutValue("Series1");
            sheet.Cells["B2"].PutValue(4);
            sheet.Cells["B3"].PutValue(2);
            sheet.Cells["B4"].PutValue(5);

            // Add a radar chart
            int chartIndex = sheet.Charts.Add(ChartType.Radar, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart.
            // isVertical = false transposes the range (plots by row) which is required for this example.
            chart.SetChartDataRange("A1:B4", false);

            // Access the first (and only) series
            Series series = chart.NSeries[0];

            // Enable data labels and show the values
            series.DataLabels.ShowValue = true;

            // Position the data labels (optional, here centered on the points)
            series.DataLabels.Position = LabelPositionType.Center;

            // For radar charts, enable category axis labels
            series.HasRadarAxisLabels = true;

            // Ensure the data label shape auto‑fits the text (fit shapes)
            series.DataLabels.IsResizeShapeToFitText = true;

            // Recalculate the chart layout before saving
            chart.Calculate();

            // Save the workbook
            workbook.Save("RadarChartWithDataLabels.xlsx");
        }
    }
}