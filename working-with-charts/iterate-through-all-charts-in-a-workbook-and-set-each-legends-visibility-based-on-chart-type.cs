// Title: Aspose.Cells C# – Loop All Workbook Charts and Control Legend Visibility by Chart Type
// Description: Shows how to create a workbook, add column and pie charts, then walk through every worksheet and each chart to set the ShowLegend property. Legends are hidden for pie‑related types (Pie, Pie3D, Doughnut, DoughnutExploded, PieExploded) and kept visible for all other chart types before saving the file.
// Keywords: Aspose.Cells chart legend | C# ShowLegend property | hide pie chart legend Aspose.Cells | iterate charts workbook .NET | chart type conditional formatting Aspose.Cells | Aspose.Cells ChartType enumeration | programmatic legend visibility | Excel chart automation C#
// Common Searches: Aspose.Cells hide legend for pie chart C# | loop through all charts in a workbook Aspose.Cells | set ShowLegend property based on ChartType Aspose.Cells | C# example toggle chart legends Aspose.Cells | Aspose.Cells chart legend visibility tutorial
// Developer Intent: Programmatically set each chart's ShowLegend flag according to its ChartType in a .NET workbook.
// Use Cases: Generate a multi‑chart report where only non‑pie charts display legends before distribution. | Load an existing template and automatically hide legends for pie, doughnut, and exploded charts during batch processing. | Standardize workbook appearance by ensuring legends are visible for column, line, or area charts while suppressing them for pie‑related charts.
// AI Prompts: Write C# code using Aspose.Cells that iterates over all worksheets and hides legends for Pie, Pie3D, Doughnut, DoughnutExploded, and PieExploded charts, leaving legends visible for other types. | Provide a complete Aspose.Cells example that adds several chart types, updates each chart's ShowLegend property based on its ChartType, and saves the workbook. | Explain how to change the legend position after toggling its visibility for each chart in an Aspose.Cells workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartLegendVisibility
{
    // Shows how to create a workbook, add column and pie charts, then walk through every worksheet and each chart to set the ShowLegend property. Legends are hidden for pie‑related types (Pie, Pie3D, Doughnut, DoughnutExploded, PieExploded) and kept visible for all other chart types before saving the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook(); // create

            // -------------------------------------------------
            // Sample data and charts for demonstration purposes
            // -------------------------------------------------
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Data";

            // Populate some data
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a Column chart (legend should be visible)
            int colChartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart colChart = sheet.Charts[colChartIdx];
            colChart.SetChartDataRange("A1:B4", true);

            // Add a Pie chart (legend will be hidden)
            int pieChartIdx = sheet.Charts.Add(ChartType.Pie, 16, 0, 26, 5);
            Chart pieChart = sheet.Charts[pieChartIdx];
            pieChart.SetChartDataRange("A1:B4", true);

            // -------------------------------------------------
            // Iterate through all worksheets and their charts
            // -------------------------------------------------
            foreach (Worksheet ws in workbook.Worksheets)
            {
                foreach (Chart chart in ws.Charts)
                {
                    // Determine visibility based on chart type
                    // Hide legend for pie‑type charts, show for others
                    bool hideLegend = chart.Type == ChartType.Pie ||
                                      chart.Type == ChartType.Pie3D ||
                                      chart.Type == ChartType.Doughnut ||
                                      chart.Type == ChartType.DoughnutExploded ||
                                      chart.Type == ChartType.PieExploded;

                    chart.ShowLegend = !hideLegend; // set visibility
                }
            }

            // Save the workbook
            workbook.Save("ChartLegendVisibilityDemo.xlsx"); // save
        }
    }
}
