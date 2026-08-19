// Title: C# Aspose.Cells – Remove All Trendlines from a Chart Before Adding New Ones
// Description: Learn how to clear every trendline from an Aspose.Cells chart using the TrendLines.Clear() method, then add a new analytical trendline (e.g., exponential) with equation and R‑squared displayed, all in a concise C# example.
// Keywords: Aspose.Cells | C# | chart trendlines | remove trendlines | clear trendlines | TrendLines.Clear | Aspose.Cells chart | add trendline | exponential trendline | trendline equation | R-squared | Aspose.Cells example | GitHub | source code
// Common Searches: Aspose.Cells remove trendlines C# | clear chart trendlines Aspose.Cells .NET | how to delete trendlines from chart using Aspose.Cells | replace chart trendline Aspose.Cells | add exponential trendline after clearing Aspose.Cells
// Developer Intent: The developer needs to delete any existing trendlines from a chart before inserting new analytical trendlines.
// Use Cases: Refresh a chart’s analytical lines after data updates by clearing old trendlines and adding updated ones. | Switch from a linear to an exponential trendline in an automated report generated with Aspose.Cells. | Prepare a clean workbook for client distribution, ensuring only the intended trendlines are present. | Reuse a chart template programmatically while guaranteeing no residual trendlines remain.
// AI Prompts: Write C# code that iterates through all series in an Aspose.Cells chart and removes each trendline using TrendLines.Clear(). | Show how to add a polynomial (order 3) trendline after clearing existing trendlines in an Aspose.Cells line chart. | Explain how to verify that a chart contains no trendlines before saving the workbook with Aspose.Cells.

using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

// Learn how to clear every trendline from an Aspose.Cells chart using the TrendLines.Clear() method, then add a new analytical trendline (e.g., exponential) with equation and R‑squared displayed, all in a concise C# example.
class RemoveTrendlinesExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data
        sheet.Cells["A1"].PutValue("X");
        sheet.Cells["B1"].PutValue("Y");
        for (int i = 2; i <= 6; i++)
        {
            sheet.Cells[$"A{i}"].PutValue(i - 1);          // X values
            sheet.Cells[$"B{i}"].PutValue((i - 1) * 2 + 1); // Y values
        }

        // Add a line chart and set its data source
        int chartIdx = sheet.Charts.Add(ChartType.Line, 5, 0, 20, 10);
        Chart chart = sheet.Charts[chartIdx];
        chart.NSeries.Add("B2:B6", true);
        chart.NSeries.CategoryData = "A2:A6";

        // Add an initial trendline to simulate existing ones
        int existingIdx = chart.NSeries[0].TrendLines.Add(TrendlineType.Linear);
        Trendline existingTrend = chart.NSeries[0].TrendLines[existingIdx];
        existingTrend.Color = Color.Blue;

        // ----------------------------------------------------
        // Remove all existing trendlines from every series
        // ----------------------------------------------------
        foreach (Series series in chart.NSeries)
        {
            series.TrendLines.Clear();
        }

        // Add new analytical trendline(s) after clearing old ones
        int newIdx = chart.NSeries[0].TrendLines.Add(TrendlineType.Exponential);
        Trendline newTrend = chart.NSeries[0].TrendLines[newIdx];
        newTrend.DisplayEquation = true;
        newTrend.DisplayRSquared = true;
        newTrend.Color = Color.Red;

        // Save the workbook
        workbook.Save("ChartWithoutOldTrendlines.xlsx");
    }
}
