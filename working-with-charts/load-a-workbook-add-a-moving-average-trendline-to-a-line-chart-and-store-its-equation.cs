using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

class MovingAverageTrendlineExample
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet (or create one if needed)
        Worksheet sheet = workbook.Worksheets[0];

        // Ensure there is data for the chart.
        // This example assumes data is already present in columns A and B.
        // If not, you can uncomment the following lines to add sample data:
        /*
        sheet.Cells["A1"].PutValue("X");
        sheet.Cells["B1"].PutValue("Y");
        for (int i = 2; i <= 10; i++)
        {
            sheet.Cells[$"A{i}"].PutValue(i - 1);          // X values
            sheet.Cells[$"B{i}"].PutValue((i - 1) * 2);   // Y values
        }
        */

        // Add a line chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Line, 5, 0, 20, 12);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data range for the series (Y values) and category (X values)
        chart.NSeries.Add("B2:B10", true);          // Y values
        chart.NSeries.CategoryData = "A2:A10";      // X values

        // Add a moving-average trendline to the first series
        int trendlineIdx = chart.NSeries[0].TrendLines.Add(TrendlineType.MovingAverage);
        Trendline trendline = chart.NSeries[0].TrendLines[trendlineIdx];

        // Configure the trendline
        trendline.Period = 3;                       // 3‑point moving average
        trendline.Name = "3‑Period Moving Average";
        trendline.DisplayEquation = true;           // Show the equation on the chart
        trendline.DisplayRSquared = false;          // R‑squared not required
        trendline.Color = Color.Red;                // Optional visual styling

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}