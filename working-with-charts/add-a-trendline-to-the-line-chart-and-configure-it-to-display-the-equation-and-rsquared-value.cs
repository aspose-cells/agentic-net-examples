using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

namespace AsposeCellsTrendlineExample
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data for the line chart
            worksheet.Cells["A1"].PutValue(1);
            worksheet.Cells["A2"].PutValue(2);
            worksheet.Cells["A3"].PutValue(3);
            worksheet.Cells["A4"].PutValue(4);
            worksheet.Cells["B1"].PutValue(2);
            worksheet.Cells["B2"].PutValue(4);
            worksheet.Cells["B3"].PutValue(6);
            worksheet.Cells["B4"].PutValue(8);

            // Add a line chart to the worksheet
            int chartIndex = worksheet.Charts.Add(ChartType.Line, 5, 0, 20, 10);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the data source for the series (Y values) and categories (X values)
            chart.NSeries.Add("B1:B4", true);
            chart.NSeries.CategoryData = "A1:A4";

            // Add a linear trendline to the first series
            int trendlineIdx = chart.NSeries[0].TrendLines.Add(TrendlineType.Linear);
            Trendline trendline = chart.NSeries[0].TrendLines[trendlineIdx];

            // Configure the trendline to display the equation and R‑squared value
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