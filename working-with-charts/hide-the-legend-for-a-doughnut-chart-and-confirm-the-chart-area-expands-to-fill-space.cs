// Title: Hide Legend in a Doughnut Chart and Auto‑Expand Chart Area with Aspose.Cells for .NET
// Description: Creates a workbook, adds sample data, inserts a doughnut chart, disables its legend, and saves the file so the plot area automatically fills the space previously occupied by the legend.
// Keywords: Aspose.Cells hide legend | doughnut chart legend removal .NET | auto expand chart area Aspose | chart layout customization C# | Aspose.Cells doughnut chart example
// Common Searches: how to hide legend in a doughnut chart using Aspose.Cells | does removing legend enlarge chart area in Excel with Aspose | Aspose.Cells ShowLegend false doughnut chart | adjust chart size after hiding legend Aspose.Cells C# | programmatically remove chart legend in .NET
// Developer Intent: Disable the legend of a doughnut chart so the chart area expands to occupy the freed space.
// Use Cases: Generate clean‑looking doughnut charts for dashboards where a legend is redundant. | Create multiple charts in a report with legends turned off, letting each chart use the full allocated region. | Validate that the plot area grows after setting ShowLegend to false for dynamic layout adjustments.
// AI Prompts: Write C# code with Aspose.Cells that hides a doughnut chart legend and verifies the plot area expands. | Show how to compare chart dimensions before and after disabling the legend in Aspose.Cells. | Explain steps to reposition and resize a doughnut chart after removing its legend using Aspose.Cells for .NET.

using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook, adds sample data, inserts a doughnut chart, disables its legend, and saves the file so the plot area automatically fills the space previously occupied by the legend.
class HideLegendDoughnutChart
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the doughnut chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("Apple");
        sheet.Cells["A3"].PutValue("Orange");
        sheet.Cells["A4"].PutValue("Banana");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(50);
        sheet.Cells["B3"].PutValue(30);
        sheet.Cells["B4"].PutValue(20);

        // Add a doughnut chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Doughnut, 5, 0, 20, 8);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data range for the chart
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Hide the legend so the chart area can expand to fill the space
        chart.ShowLegend = false;

        // Save the workbook with the modified chart
        workbook.Save("DoughnutChart_NoLegend.xlsx");
    }
}
