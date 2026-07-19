// Title: Hide Chart Legend in Aspose.Cells for .NET to Maximize Plot Area
// Description: Shows how to create a workbook, add a column chart, and remove its legend by setting Chart.ShowLegend = false, then save the file as ChartWithoutLegend.xlsx.
// Keywords: Aspose.Cells hide legend | Aspose.Cells chart legend removal | C# Chart.ShowLegend false | increase chart plotting area | Aspose.Cells chart formatting .NET | remove legend Excel chart Aspose | chart layout optimization | Aspose.Cells example C#
// Common Searches: Aspose.Cells hide legend C# | remove legend from chart Aspose.Cells .NET | Chart.ShowLegend property usage | enlarge chart area Aspose.Cells | Aspose.Cells chart formatting guide
// Developer Intent: Hide the chart legend to free space and enlarge the plotting region.
// Use Cases: Compact column chart for a dashboard where the legend adds no value. | Excel report generation where series names are self‑explanatory. | Slide‑deck chart where the legend would overlap data points.
// AI Prompts: Generate code to hide the legend for any Aspose.Cells chart type in C#. | Provide a snippet that toggles chart legend visibility based on a runtime condition. | Explain how to adjust the chart area after removing the legend to fully utilize the plot space.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace RemoveLegendExample
{
    // Shows how to create a workbook, add a column chart, and remove its legend by setting Chart.ShowLegend = false, then save the file as ChartWithoutLegend.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
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

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Hide the legend to maximize the plotting area
            chart.ShowLegend = false;

            // Save the workbook
            workbook.Save("ChartWithoutLegend.xlsx");
        }
    }
}
