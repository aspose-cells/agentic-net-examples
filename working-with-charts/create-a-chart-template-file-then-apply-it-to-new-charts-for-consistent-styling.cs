// Title: Create a reusable chart template workbook and apply its styling to new column charts using Aspose.Cells for .NET
// AI Prompts: Generate C# code that builds a styled column chart as a template, saves it to a workbook, and then reuses the template's formatting for another chart in a separate workbook with Aspose.Cells. | Show how to programmatically copy a chart's title, font color, plot‑area background, and legend position from a template chart to a new chart that uses different data series in Aspose.Cells. | Provide a .NET example that creates a chart template, binds a new data series to a fresh chart, transfers the template's visual settings, and saves the final workbook.
// Common Searches: aspnet how to create a chart styling template workbook with Aspose.Cells and reuse it | copy chart title and legend formatting from one workbook to another using C# Aspose.Cells | apply same styling to multiple column charts in Aspose.Cells .NET | example of reusing chart template for different data series in Aspose.Cells
// Tags: Aspose.Cells chart template styling | duplicate chart formatting C# | reuse chart visual settings Aspose.Cells | column chart template workbook | programmatic chart theme application .NET

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook containing a styled column chart that serves as a template, saves it, then opens a second workbook, adds a new column chart, copies the template's title, font color, plot‑area background, and legend position, binds a new data series, and saves the final workbook with the transferred styling.
class ChartTemplateExample
{
    static void Main()
    {
        try
        {
            // -------------------------------------------------
            // 1. Create a workbook that will hold the chart template
            // -------------------------------------------------
            Workbook templateWb = new Workbook();
            Worksheet templateWs = templateWb.Worksheets[0];

            // Sample data for the template chart
            templateWs.Cells["A1"].PutValue("Month");
            templateWs.Cells["B1"].PutValue("Sales");
            templateWs.Cells["A2"].PutValue("Jan");
            templateWs.Cells["A3"].PutValue("Feb");
            templateWs.Cells["A4"].PutValue("Mar");
            templateWs.Cells["B2"].PutValue(120);
            templateWs.Cells["B3"].PutValue(150);
            templateWs.Cells["B4"].PutValue(180);

            // Add a column chart that will become the template
            int tmplChartIdx = templateWs.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart tmplChart = templateWs.Charts[tmplChartIdx];

            // Styling the template chart
            tmplChart.Title.Text = "Quarterly Sales";
            tmplChart.Title.Font.Color = Color.Blue;
            tmplChart.PlotArea.Area.ForegroundColor = Color.LightYellow;
            tmplChart.Legend.Position = LegendPositionType.Right;

            // Add a series (placeholder) to the template
            int seriesIdx = tmplChart.NSeries.Add("B2:B4", true);
            tmplChart.NSeries[seriesIdx].Name = "Sales";

            // Save the workbook that contains the template chart (optional)
            templateWb.Save("TemplateWorkbook.xlsx");

            // -------------------------------------------------
            // 2. Create a new workbook and apply the saved template styling manually
            // -------------------------------------------------
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];

            // Data for the new chart
            ws.Cells["A1"].PutValue("Month");
            ws.Cells["B1"].PutValue("Revenue");
            ws.Cells["A2"].PutValue("Apr");
            ws.Cells["A3"].PutValue("May");
            ws.Cells["A4"].PutValue("Jun");
            ws.Cells["B2"].PutValue(200);
            ws.Cells["B3"].PutValue(250);
            ws.Cells["B4"].PutValue(300);

            // Add an empty chart where the template styling will be applied
            int chartIdx = ws.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = ws.Charts[chartIdx];

            // Apply the same styling as the template chart
            chart.Title.Text = tmplChart.Title.Text;
            chart.Title.Font.Color = tmplChart.Title.Font.Color;
            chart.PlotArea.Area.ForegroundColor = tmplChart.PlotArea.Area.ForegroundColor;
            chart.Legend.Position = tmplChart.Legend.Position;

            // Bind the new data to the series (category + values)
            int newSeriesIdx = chart.NSeries.Add("A2:A4", true); // category data
            chart.NSeries[newSeriesIdx].Values = "B2:B4";
            chart.NSeries[newSeriesIdx].Name = "Revenue";

            // Save the final workbook containing the styled chart
            wb.Save("WorkbookWithStyledChart.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
