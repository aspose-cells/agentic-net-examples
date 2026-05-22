using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

class RemoveTrendlinesExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        sheet.Cells["A1"].PutValue("X");
        sheet.Cells["B1"].PutValue("Y");
        for (int i = 2; i <= 6; i++)
        {
            sheet.Cells[$"A{i}"].PutValue(i - 1);          // X values
            sheet.Cells[$"B{i}"].PutValue((i - 1) * 2);   // Y values
        }

        // Add a line chart and set its data source
        int chartIdx = sheet.Charts.Add(ChartType.Line, 5, 0, 20, 10);
        Chart chart = sheet.Charts[chartIdx];
        chart.NSeries.Add("B2:B6", true);
        chart.NSeries.CategoryData = "A2:A6";

        // Add an initial trendline (to demonstrate that it will be removed)
        int oldTrendIdx = chart.NSeries[0].TrendLines.Add(TrendlineType.Linear);
        chart.NSeries[0].TrendLines[oldTrendIdx].Name = "Old Trend";

        // ----------------------------------------------------
        // Remove all existing trendlines from every series
        // ----------------------------------------------------
        foreach (Series series in chart.NSeries)
        {
            // Clear the TrendLines collection for the current series
            series.TrendLines.Clear();
        }

        // Add new analytical trendlines after clearing old ones
        int newTrendIdx = chart.NSeries[0].TrendLines.Add(TrendlineType.Exponential, "New Exponential");
        Trendline newTrend = chart.NSeries[0].TrendLines[newTrendIdx];
        newTrend.DisplayEquation = true;
        newTrend.DisplayRSquared = true;
        newTrend.Color = Color.Red;

        // Save the workbook
        workbook.Save("ChartWithoutOldTrendlines.xlsx");
    }
}