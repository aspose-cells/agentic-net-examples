// Title: Hide Legend for Single‑Series Charts with Aspose.Cells for .NET
// Description: Creates an Excel workbook with two column charts, checks each chart's series count, disables the legend for charts that contain only one series, keeps it visible for multi‑series charts, and saves the file using Aspose.Cells for C#.
// Keywords: Aspose.Cells C# chart legend | hide legend single series Aspose.Cells | ShowLegend property .NET | conditional chart legend Excel | Aspose.Cells chart automation | Excel single series legend hide | Aspose.Cells conditional formatting | C# Excel chart legend control
// Common Searches: Aspose.Cells hide legend for single series chart | C# set chart legend visibility based on series count | ShowLegend property Aspose.Cells example | How to remove redundant legend in Excel chart using Aspose | Conditional legend display Aspose.Cells .NET
// Developer Intent: Programmatically hide the legend on charts that have only one data series while ensuring legends remain visible on charts with multiple series.
// Use Cases: Generating Excel reports where single‑series charts should not display a legend to reduce visual clutter. | Building dashboards that mix single‑ and multi‑series charts and need automatic legend management. | Exporting analytical data to Excel and ensuring legends appear only when they add value.
// AI Prompts: Provide C# code using Aspose.Cells that iterates over all worksheet charts and sets ShowLegend to false when the chart has exactly one series. | Show an example of conditional legend visibility for Excel charts in Aspose.Cells, including both single‑series and multi‑series scenarios. | Explain how the ShowLegend property works in Aspose.Cells and how to use it to hide redundant legends in generated Excel files.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace HideLegendForSingleSeriesChart
{
    // Creates an Excel workbook with two column charts, checks each chart's series count, disables the legend for charts that contain only one series, keeps it visible for multi‑series charts, and saves the file using Aspose.Cells for C#.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Sample data for the first chart (single series)
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a chart with a single series
            int singleSeriesChartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart singleSeriesChart = sheet.Charts[singleSeriesChartIdx];
            singleSeriesChart.NSeries.Add("B2:B4", true);
            singleSeriesChart.NSeries.CategoryData = "A2:A4";

            // Sample data for the second chart (multiple series)
            sheet.Cells["C1"].PutValue("Series1");
            sheet.Cells["C2"].PutValue(15);
            sheet.Cells["C3"].PutValue(25);
            sheet.Cells["C4"].PutValue(35);
            sheet.Cells["D1"].PutValue("Series2");
            sheet.Cells["D2"].PutValue(12);
            sheet.Cells["D3"].PutValue(22);
            sheet.Cells["D4"].PutValue(32);

            // Add a chart with two series
            int multiSeriesChartIdx = sheet.Charts.Add(ChartType.Column, 20, 0, 30, 5);
            Chart multiSeriesChart = sheet.Charts[multiSeriesChartIdx];
            multiSeriesChart.NSeries.Add("C2:C4", true);
            multiSeriesChart.NSeries.Add("D2:D4", true);
            multiSeriesChart.NSeries.CategoryData = "A2:A4";

            // Iterate through all charts in the worksheet
            foreach (Chart chart in sheet.Charts)
            {
                // If the chart has only one series, hide its legend
                if (chart.NSeries.Count == 1)
                {
                    chart.ShowLegend = false;
                }
                else
                {
                    // Ensure legend is visible for charts with multiple series
                    chart.ShowLegend = true;
                }
            }

            // Save the workbook
            workbook.Save("ChartsWithConditionalLegend.xlsx");
        }
    }
}
