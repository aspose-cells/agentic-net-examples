// Title: Hide Chart Legend in an Aspose.Cells .NET Dashboard Worksheet
// Description: Demonstrates how to create a dashboard sheet with a column chart in Aspose.Cells for .NET, assign data ranges, and remove the legend (chart.ShowLegend = false) to maximize the chart's display area before saving the workbook.
// Keywords: Aspose.Cells hide legend | C# chart legend removal | Aspose.Cells dashboard chart | chart.ShowLegend false | maximize chart area .NET | Aspose.Cells chart formatting | Excel dashboard Aspose.Cells
// Common Searches: how to hide chart legend using Aspose.Cells C# | remove legend from Aspose.Cells chart on dashboard sheet | Aspose.Cells .NET increase chart size by hiding legend | Aspose.Cells chart.ShowLegend property example | C# hide legend in Excel chart programmatically
// Developer Intent: Programmatically hide a chart legend on a dashboard worksheet to free up visual space.
// Use Cases: Design a sales dashboard where the legend is omitted to give the chart more room. | Generate KPI reports with multiple charts and suppress legends to reduce visual clutter. | Prepare workbooks for PDF export by disabling legends on all charts to improve layout.
// AI Prompts: Provide C# code that adds a line chart with Aspose.Cells and hides its legend. | Show how to loop through every chart in a worksheet and set ShowLegend = false. | Explain how to conditionally hide a chart legend based on the number of series in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsDashboard
{
    // Demonstrates how to create a dashboard sheet with a column chart in Aspose.Cells for .NET, assign data ranges, and remove the legend (chart.ShowLegend = false) to maximize the chart's display area before saving the workbook.
    class HideChartLegend
    {
        static void Main()
        {
            // Create a new workbook (dashboard sheet)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Dashboard";

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Q1");
            sheet.Cells["A3"].PutValue("Q2");
            sheet.Cells["A4"].PutValue("Q3");
            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["B2"].PutValue(15000);
            sheet.Cells["B3"].PutValue(23000);
            sheet.Cells["B4"].PutValue(18000);

            // Add a column chart to the dashboard sheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Hide the legend to maximize the chart display area
            chart.ShowLegend = false;

            // Save the workbook
            workbook.Save("DashboardWithHiddenLegend.xlsx");
        }
    }
}
