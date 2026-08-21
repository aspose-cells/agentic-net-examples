// Title: Hide Legend for Single‑Series Charts with Aspose.Cells for .NET (C#)
// Description: This C# example demonstrates how to programmatically hide the legend of a chart that contains only one data series using Aspose.Cells. The code creates a workbook with two column charts, iterates through every worksheet and chart, sets ShowLegend to false when the series count equals one, keeps the legend visible for multi‑series charts, and saves the file as an Excel workbook.
// Keywords: Aspose.Cells | C# chart legend | hide legend single series | ShowLegend property | conditional legend visibility | .NET Excel automation | remove redundant legend | Excel chart customization
// Common Searches: Aspose.Cells hide chart legend single series | C# set ShowLegend false for one series chart | how to hide legend in Aspose.Cells chart | conditional legend visibility Aspose.Cells .NET | remove legend from Excel chart with one series
// Developer Intent: Automatically suppress the legend on charts that have only one data series while preserving it on charts with multiple series.
// Use Cases: Generating financial reports where single‑series column charts should not display a legend. | Building dashboards that contain many charts and need consistent legend handling based on series count. | Applying a workbook‑wide rule to ensure legends appear only when they add value to the visualisation.
// AI Prompts: Provide C# code using Aspose.Cells to hide the legend of any chart that contains exactly one series. | Show an Aspose.Cells snippet that toggles ShowLegend based on the number of series for each chart in a workbook. | Explain how to iterate through all worksheets and set legend visibility only for charts with more than one series in Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// This C# example demonstrates how to programmatically hide the legend of a chart that contains only one data series using Aspose.Cells. The code creates a workbook with two column charts, iterates through every worksheet and chart, sets ShowLegend to false when the series count equals one, keeps the legend visible for multi‑series charts, and saves the file as an Excel workbook.
class HideLegendSingleSeries
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // -----------------------------
        // Chart with a single series
        // -----------------------------
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["A4"].PutValue("C");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["B4"].PutValue(30);
        int chartIdx1 = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
        Chart chart1 = sheet.Charts[chartIdx1];
        chart1.NSeries.Add("B2:B4", true);
        chart1.NSeries.CategoryData = "A2:A4";

        // -----------------------------
        // Chart with multiple series
        // -----------------------------
        sheet.Cells["D1"].PutValue("Category");
        sheet.Cells["D2"].PutValue("A");
        sheet.Cells["D3"].PutValue("B");
        sheet.Cells["D4"].PutValue("C");
        sheet.Cells["E1"].PutValue("Series1");
        sheet.Cells["E2"].PutValue(5);
        sheet.Cells["E3"].PutValue(15);
        sheet.Cells["E4"].PutValue(25);
        sheet.Cells["F1"].PutValue("Series2");
        sheet.Cells["F2"].PutValue(8);
        sheet.Cells["F3"].PutValue(18);
        sheet.Cells["F4"].PutValue(28);
        int chartIdx2 = sheet.Charts.Add(ChartType.Column, 20, 0, 30, 5);
        Chart chart2 = sheet.Charts[chartIdx2];
        chart2.NSeries.Add("E2:E4", true);
        chart2.NSeries.Add("F2:F4", true);
        chart2.NSeries.CategoryData = "D2:D4";

        // Iterate through all worksheets and their charts
        foreach (Worksheet ws in workbook.Worksheets)
        {
            foreach (Chart ch in ws.Charts)
            {
                // Hide legend if the chart contains only one series
                if (ch.NSeries.Count == 1)
                {
                    ch.ShowLegend = false;
                }
                else
                {
                    // Ensure legend is visible for charts with multiple series
                    ch.ShowLegend = true;
                }
            }
        }

        // Save the workbook
        workbook.Save("HideLegendSingleSeries.xlsx");
    }
}
