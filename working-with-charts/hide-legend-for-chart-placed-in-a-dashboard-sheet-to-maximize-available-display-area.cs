// Title: Hide Chart Legend in a Dashboard Sheet with Aspose.Cells for .NET (C#)
// Description: C# example that creates a workbook, adds a "Dashboard" worksheet, populates sample data, inserts a column chart, disables the legend via Chart.ShowLegend = false, and saves the file as DashboardChart_NoLegend.xlsx.
// Keywords: Aspose.Cells hide chart legend | Chart.ShowLegend false C# | remove legend Aspose.Cells .NET | dashboard chart without legend | Aspose.Cells chart customization | C# Aspose.Cells example | Excel chart legend visibility | Aspose.Cells GitHub sample
// Common Searches: how to hide legend in Aspose.Cells chart C# | Aspose.Cells Chart.ShowLegend property example | remove chart legend to maximize area Aspose.Cells | dashboard sheet chart without legend Aspose | Aspose.Cells hide legend code snippet
// Developer Intent: The developer needs to suppress the chart legend so the chart occupies the full available space on a dashboard worksheet.
// Use Cases: Design compact dashboards where legends would waste space. | Generate reports with multiple charts that share a common legend elsewhere. | Create printable workbooks with a clean visual layout by omitting legends.
// AI Prompts: Generate C# code to hide the legend for any Aspose.Cells chart type. | Show how to toggle Chart.ShowLegend based on a runtime condition in Aspose.Cells. | Explain the impact of Chart.ShowLegend on chart layout and other formatting options in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsDashboardExample
{
    // C# example that creates a workbook, adds a "Dashboard" worksheet, populates sample data, inserts a column chart, disables the legend via Chart.ShowLegend = false, and saves the file as DashboardChart_NoLegend.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle create)
            Workbook workbook = new Workbook();

            // Add a worksheet that will act as a dashboard
            Worksheet dashboard = workbook.Worksheets[0];
            dashboard.Name = "Dashboard";

            // Populate some sample data for the chart
            dashboard.Cells["A1"].PutValue("Category");
            dashboard.Cells["A2"].PutValue("A");
            dashboard.Cells["A3"].PutValue("B");
            dashboard.Cells["A4"].PutValue("C");
            dashboard.Cells["B1"].PutValue("Value");
            dashboard.Cells["B2"].PutValue(10);
            dashboard.Cells["B3"].PutValue(20);
            dashboard.Cells["B4"].PutValue(30);

            // Add a column chart to the dashboard sheet
            int chartIndex = dashboard.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = dashboard.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Hide the legend to maximize the chart display area
            chart.ShowLegend = false; // using Chart.ShowLegend property

            // Save the workbook (lifecycle save)
            workbook.Save("DashboardChart_NoLegend.xlsx");
        }
    }
}
