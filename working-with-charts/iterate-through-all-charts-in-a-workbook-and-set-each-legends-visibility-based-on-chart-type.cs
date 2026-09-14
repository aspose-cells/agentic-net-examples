// Title: Iterate over all charts in a worksheet and toggle legend visibility by chart type using Aspose.Cells for .NET
// AI Prompts: Write C# code that loops through every chart in a workbook and disables the legend for pie and 3‑D pie charts while keeping it enabled for all other chart types with Aspose.Cells. | Create a sample that adds a column chart and a pie chart to a worksheet, then uses the Chart.ShowLegend property to set legend visibility conditionally for each chart.
// Common Searches: Aspose.Cells C# hide legend for pie chart programmatically | how to set chart legend visibility based on chart type in Aspose.Cells | iterate all charts in an Excel workbook and change legend display .NET | Chart.ShowLegend example for conditional formatting Aspose.Cells | C# code to show legends only on non‑pie charts using Aspose.Cells
// Tags: Chart.ShowLegend property Aspose.Cells | iterate worksheet charts C# | hide legend for pie charts Aspose.Cells | conditional legend visibility by chart type | apply chart formatting in .NET workbook

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The program creates sample data, adds a column chart and a pie chart, then iterates through all charts on the first worksheet, turning off the legend for pie and 3‑D pie charts while keeping it on for other chart types, and finally saves the workbook.
class Program
{
    static void Main()
    {
        // Create a new workbook (creation rule)
        Workbook workbook = new Workbook();

        // Prepare sample data
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["A4"].PutValue("C");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["B4"].PutValue(30);

        // Add a column chart
        int columnChartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
        Chart columnChart = sheet.Charts[columnChartIdx];
        columnChart.NSeries.Add("B2:B4", true);
        columnChart.NSeries.CategoryData = "A2:A4";

        // Add a pie chart
        int pieChartIdx = sheet.Charts.Add(ChartType.Pie, 16, 0, 26, 5);
        Chart pieChart = sheet.Charts[pieChartIdx];
        pieChart.NSeries.Add("B2:B4", true);
        pieChart.NSeries.CategoryData = "A2:A4";

        // Iterate through all charts in the first worksheet
        foreach (Chart chart in sheet.Charts)
        {
            // Hide legend for pie charts (including 3‑D pie), show for others
            if (chart.Type == ChartType.Pie || chart.Type == ChartType.Pie3D)
                chart.ShowLegend = false; // use Chart.ShowLegend property
            else
                chart.ShowLegend = true;
        }

        // Save the workbook (save rule)
        workbook.Save("ChartsLegendVisibility.xlsx");
    }
}
