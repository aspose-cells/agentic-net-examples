// Title: Hide the legend of a doughnut chart and let the chart area automatically expand with Aspose.Cells for .NET
// AI Prompts: Generate a new workbook, insert sample data, create a doughnut chart, set ShowLegend to false, and save as an .xlsx file using Aspose.Cells in C#. | Locate an existing doughnut chart in a worksheet, set its ShowLegend property to false so the chart area fills the space, and print the legend state to the console. | Write C# code that adds a doughnut chart, disables the legend display, and verifies the chart layout expands automatically with Aspose.Cells.
// Common Searches: Aspose.Cells C# hide legend on doughnut chart and expand chart area | How to remove legend from a doughnut chart using Aspose.Cells for .NET | Chart.ShowLegend false effect on layout in Aspose.Cells | Resize doughnut chart automatically after hiding legend in C# | Programmatically hide legend in Excel doughnut chart with Aspose.Cells
// Tags: Aspose.Cells hide chart legend | doughnut chart layout adjustment Aspose.Cells | chart.ShowLegend property C# | auto expand chart area after legend removal | Aspose.Cells generate doughnut chart without legend

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, adds sample data, builds a doughnut chart, disables its legend by setting ShowLegend to false (causing the chart area to expand), prints the legend visibility status, and saves the file as DoughnutChart_NoLegend.xlsx.
class HideDoughnutLegend
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

        // Hide the legend; the chart area will automatically expand to use the freed space
        chart.ShowLegend = false;

        // Verify that the legend is hidden (optional)
        Console.WriteLine("Legend visible? " + chart.ShowLegend);

        // Save the workbook
        workbook.Save("DoughnutChart_NoLegend.xlsx");
    }
}
