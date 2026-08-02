// Title: Hide Chart Legend in a Dashboard Widget Using Aspose.Cells for .NET (C#)
// Description: The sample builds a workbook, fills cells A1:B4 with categories and values, inserts a column chart positioned as a dashboard widget, disables the legend via the ShowLegend property, and saves the file as DashboardChart_NoLegend.xlsx.
// Keywords: Aspose.Cells | C# | chart legend removal | ShowLegend property | dashboard chart | column chart | Excel automation | maximize chart area | Aspose.Cells .NET
// Common Searches: Aspose.Cells hide legend C# | ShowLegend false example Aspose | dashboard chart without legend .NET | remove chart legend programmatically | increase chart display area Aspose.Cells
// Developer Intent: Eliminate the legend of a chart placed on a dashboard sheet to maximize usable space.
// Use Cases: Create a compact column chart for a KPI dashboard where the legend adds no value. | Design multiple dashboard widgets on a single worksheet while preserving screen real estate. | Generate Excel reports with clean‑looking charts that omit legends for a streamlined presentation.
// AI Prompts: Write C# code with Aspose.Cells that hides a chart legend and resizes the chart for a dashboard layout. | Show how to loop through all charts in a workbook and turn off their legends using Aspose.Cells for .NET. | Explain the effect of the ShowLegend property and when it should be disabled based on chart type.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsDashboardChart
{
    // The sample builds a workbook, fills cells A1:B4 with categories and values, inserts a column chart positioned as a dashboard widget, disables the legend via the ShowLegend property, and saves the file as DashboardChart_NoLegend.xlsx.
    class HideLegendDemo
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a chart to the worksheet (positioned like a dashboard widget)
            int chartIndex = sheet.Charts.Add(ChartType.Column, 2, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Hide the legend to maximize the display area
            chart.ShowLegend = false;

            // Save the workbook
            workbook.Save("DashboardChart_NoLegend.xlsx");
        }
    }
}
