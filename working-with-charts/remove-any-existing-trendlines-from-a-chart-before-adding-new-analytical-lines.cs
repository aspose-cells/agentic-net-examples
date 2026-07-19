// Title: How to Remove All Trendlines from an Aspose.Cells Chart and Add New Ones in C#
// Description: Demonstrates creating a workbook, inserting a line chart with sample data, deleting every existing trendline from each series, and then adding a new exponential trendline with equation and R‑squared displayed, using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# remove trendlines | clear chart trendlines Aspose.Cells | delete trendline series Aspose.Cells | add new trendline Aspose.Cells | Aspose.Cells chart API | trendline management .NET | Aspose.Cells remove all trendlines | chart series trendline collection
// Common Searches: Aspose.Cells delete all trendlines from chart | C# remove trendlines before adding new ones Aspose.Cells | clear TrendLines collection in Aspose.Cells chart | how to replace a trendline in Aspose.Cells | remove existing trendlines Aspose.Cells .NET
// Developer Intent: Remove every trendline attached to chart series, then insert updated analytical trendlines.
// Use Cases: Refresh analytical lines after data updates without duplicating old trendlines. | Prepare a clean report chart that only shows newly added trendlines. | Switch between different trendline types (linear, exponential, polynomial) by clearing previous ones first.
// AI Prompts: Generate C# code with Aspose.Cells that clears all trendlines from each series in a chart and adds a polynomial trendline with custom color and label. | Explain step‑by‑step how to iterate through chart series in Aspose.Cells and remove their TrendLines collections before adding new analytical lines. | Provide a tutorial for replacing an existing linear trendline with an exponential trendline in an Aspose.Cells chart, including formatting options.

using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

// Demonstrates creating a workbook, inserting a line chart with sample data, deleting every existing trendline from each series, and then adding a new exponential trendline with equation and R‑squared displayed, using Aspose.Cells for .NET.
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
            sheet.Cells[$"B{i}"].PutValue((i - 1) * 2 + 1); // Y values
        }

        // Add a line chart to the worksheet
        int chartIdx = sheet.Charts.Add(ChartType.Line, 5, 0, 20, 10);
        Chart chart = sheet.Charts[chartIdx];
        chart.NSeries.Add("B2:B6", true);          // Y values
        chart.NSeries.CategoryData = "A2:A6";      // X values

        // Add an initial trendline (this will be removed later)
        int oldTrendIdx = chart.NSeries[0].TrendLines.Add(TrendlineType.Linear);
        Trendline oldTrend = chart.NSeries[0].TrendLines[oldTrendIdx];
        oldTrend.Name = "Old Trend";
        oldTrend.Color = Color.Gray;

        // ---------------------------------------------------------
        // Remove all existing trendlines from each series in the chart
        // ---------------------------------------------------------
        foreach (Series series in chart.NSeries)
        {
            // Clear the TrendLines collection for the current series
            series.TrendLines.Clear();
        }

        // Add new analytical trendlines after removal
        int newTrendIdx = chart.NSeries[0].TrendLines.Add(TrendlineType.Exponential, "New Trend");
        Trendline newTrend = chart.NSeries[0].TrendLines[newTrendIdx];
        newTrend.DisplayEquation = true;
        newTrend.DisplayRSquared = true;
        newTrend.Color = Color.Red;

        // Save the workbook with the updated chart
        workbook.Save("ChartWithoutOldTrendlines.xlsx", SaveFormat.Xlsx);
    }
}
