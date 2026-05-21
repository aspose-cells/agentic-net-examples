using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

namespace AsposeCellsTrendlineExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data for the line chart
            worksheet.Cells["A1"].PutValue(10);
            worksheet.Cells["A2"].PutValue(20);
            worksheet.Cells["A3"].PutValue(30);
            worksheet.Cells["A4"].PutValue(40);
            worksheet.Cells["B1"].PutValue(15);
            worksheet.Cells["B2"].PutValue(25);
            worksheet.Cells["B3"].PutValue(35);
            worksheet.Cells["B4"].PutValue(45);

            // Add a line chart to the worksheet
            int chartIndex = worksheet.Charts.Add(ChartType.Line, 5, 0, 20, 8);
            Chart chart = worksheet.Charts[chartIndex];

            // Add the data series to the chart (Y values)
            chart.NSeries.Add("B1:B4", true);
            // Set category (X) data
            chart.NSeries.CategoryData = "A1:A4";

            // Add a linear trendline to the first series
            int trendlineIdx = chart.NSeries[0].TrendLines.Add(TrendlineType.Linear);
            Trendline trendline = chart.NSeries[0].TrendLines[trendlineIdx];

            // Configure the trendline to display equation and R‑squared value
            trendline.DisplayEquation = true;
            trendline.DisplayRSquared = true;

            // Optional: set a custom name or color for the trendline
            trendline.Name = "Linear Trend";
            trendline.Color = Color.Blue;

            // Save the workbook to a file
            workbook.Save("LineChartWithTrendline.xlsx");
        }
    }
}