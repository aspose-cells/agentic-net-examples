// Title: Hide chart legend for single‑series column charts using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that creates a workbook, adds charts, and automatically disables the legend when a chart contains only one data series. | Show how to iterate over every worksheet and chart in a workbook and set ShowLegend = false for charts that have a single series before saving the file.
// Common Searches: Aspose.Cells hide legend for chart with one series C# | C# loop through workbook charts and remove legend if only one series Aspose.Cells | How to suppress redundant legend in a single‑series column chart using Aspose.Cells .NET | Set ShowLegend property false for single‑series charts Aspose.Cells example
// Tags: Aspose.Cells chart legend suppression | ShowLegend false single series | C# iterate workbook charts Aspose.Cells | column chart single series Aspose.Cells | remove redundant legend Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace HideLegendForSingleSeriesChart
{
    // The program creates a workbook with both single‑series and multi‑series column charts, iterates through all charts, disables the legend for any chart that has only one series by setting ShowLegend to false, and saves the result as HideLegendSingleSeries.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Sample data for a chart with a single series
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a chart (single series)
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Add another chart with multiple series for demonstration
            int multiChartIdx = sheet.Charts.Add(ChartType.Column, 20, 0, 30, 5);
            Chart multiChart = sheet.Charts[multiChartIdx];
            // Second series data
            sheet.Cells["C1"].PutValue("Series2");
            sheet.Cells["C2"].PutValue(15);
            sheet.Cells["C3"].PutValue(25);
            sheet.Cells["C4"].PutValue(35);
            multiChart.NSeries.Add("B2:B4", true);
            multiChart.NSeries.Add("C2:C4", true);
            multiChart.NSeries.CategoryData = "A2:A4";

            // Iterate through all worksheets and their charts
            foreach (Worksheet ws in workbook.Worksheets)
            {
                foreach (Chart ch in ws.Charts)
                {
                    // If the chart has only one series, hide its legend
                    if (ch.NSeries.Count == 1)
                    {
                        ch.ShowLegend = false;
                    }
                }
            }

            // Save the workbook
            workbook.Save("HideLegendSingleSeries.xlsx");
        }
    }
}
