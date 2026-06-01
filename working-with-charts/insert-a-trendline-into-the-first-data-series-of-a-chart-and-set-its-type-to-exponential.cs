using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

namespace TrendlineExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data for the chart
            sheet.Cells["A1"].PutValue("X");
            sheet.Cells["B1"].PutValue("Y");
            for (int i = 2; i <= 5; i++)
            {
                sheet.Cells["A" + i].PutValue(i);          // X values
                sheet.Cells["B" + i].PutValue(i * i);      // Y values (quadratic)
            }

            // Add a line chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Line, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data series for the chart (Y values) and category (X values)
            chart.NSeries.Add("B2:B5", true);
            chart.NSeries[0].XValues = "A2:A5";

            // Insert an exponential trendline into the first data series
            int trendlineIdx = chart.NSeries[0].TrendLines.Add(TrendlineType.Exponential);
            Trendline trendline = chart.NSeries[0].TrendLines[trendlineIdx];

            // Optional: display equation and R‑squared value, and set a custom color
            trendline.DisplayEquation = true;
            trendline.DisplayRSquared = true;
            trendline.Color = Color.Red;

            // Save the workbook to a file
            workbook.Save("TrendlineExponential.xlsx");
        }
    }
}