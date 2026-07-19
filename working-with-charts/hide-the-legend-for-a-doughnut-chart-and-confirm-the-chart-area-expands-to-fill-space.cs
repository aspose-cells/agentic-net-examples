// Title: Hide Legend and Expand Chart Area for a Doughnut Chart – Aspose.Cells for .NET (C#)
// Description: Creates a workbook, adds sample data, inserts a doughnut chart, disables its legend with chart.ShowLegend = false, verifies the setting, and saves the file. Hiding the legend automatically expands the chart area to fill the freed space.
// Keywords: Aspose.Cells | C# | doughnut chart | hide legend | chart.ShowLegend | expand chart area | chart area fill | Excel chart formatting | Aspose.Cells chart example | remove chart legend
// Common Searches: Aspose.Cells hide legend doughnut chart C# | expand chart area after removing legend Aspose.Cells | chart.ShowLegend false example | create doughnut chart without legend using Aspose.Cells | C# code to hide chart legend in Excel workbook
// Developer Intent: Remove the legend from a doughnut chart so the chart area occupies the full plot region.
// Use Cases: Design compact financial dashboards where doughnut charts need maximum visual space. | Generate automated reports that omit legends for cleaner presentation of multiple charts. | Programmatically confirm legend visibility before publishing an Excel workbook to ensure consistent layout.
// AI Prompts: Generate C# code with Aspose.Cells that builds a doughnut chart, hides its legend, and lets the chart area expand to fill the space. | Show how to use chart.ShowLegend = false, verify the property, and save the workbook as an Excel file. | Explain how hiding a chart legend affects the chart area dimensions and how to adjust the chart size if needed.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook, adds sample data, inserts a doughnut chart, disables its legend with chart.ShowLegend = false, verifies the setting, and saves the file. Hiding the legend automatically expands the chart area to fill the freed space.
class HideLegendDoughnutChart
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add sample data for the doughnut chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("Apple");
        sheet.Cells["A3"].PutValue("Orange");
        sheet.Cells["A4"].PutValue("Banana");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(50);
        sheet.Cells["B3"].PutValue(30);
        sheet.Cells["B4"].PutValue(20);

        // Insert a doughnut chart
        int chartIndex = sheet.Charts.Add(ChartType.Doughnut, 5, 0, 20, 8);
        Chart chart = sheet.Charts[chartIndex];

        // Define the data range for the chart
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Hide the legend so the chart area expands to fill the space
        chart.ShowLegend = false;

        // Verify that the legend is hidden
        Console.WriteLine("Legend visible? " + chart.ShowLegend);

        // Save the workbook
        workbook.Save("DoughnutChart_NoLegend.xlsx");
    }
}
