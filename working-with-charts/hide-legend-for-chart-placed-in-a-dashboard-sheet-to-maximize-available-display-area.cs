// Title: How to hide a chart legend on a dashboard worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code with Aspose.Cells that creates a column chart on a worksheet named "Dashboard" and disables its legend. | Show the steps to set the ShowLegend property to false for a chart placed in a dashboard sheet with Aspose.Cells. | Provide a complete example that adds sample data, inserts a column chart, hides the legend, and saves the workbook as an .xlsx file.
// Common Searches: asp.net hide legend column chart Aspose.Cells dashboard sheet | Aspose.Cells C# remove chart legend to increase plot area | set ShowLegend false for chart in Aspose.Cells workbook example | how to maximize chart area by hiding legend in Aspose.Cells .NET | sample code Aspose.Cells create dashboard worksheet with chart without legend
// Tags: Aspose.Cells chart ShowLegend property | C# hide chart legend Aspose.Cells | dashboard worksheet column chart Aspose.Cells | increase chart plot area Aspose.Cells | save workbook without legend Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a new workbook, adds sample data, inserts a column chart on a worksheet named "Dashboard", hides the chart legend by setting ShowLegend to false, and saves the file as DashboardChart_NoLegend.xlsx.
class HideChartLegendDashboard
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet (used as a dashboard sheet) and give it a name
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Name = "Dashboard";

        // Populate sample data for the chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["A4"].PutValue("C");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["B4"].PutValue(30);

        // Add a column chart to the dashboard sheet
        // Parameters: chart type, upper-left row, upper-left column, lower-right row, lower-right column
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
        Chart chart = sheet.Charts[chartIndex];

        // Define the data range for the chart
        chart.NSeries.Add("B2:B4", true);          // Values
        chart.NSeries.CategoryData = "A2:A4";      // Categories

        // Hide the legend to maximize the available display area
        chart.ShowLegend = false;

        // Save the workbook with the chart that has no legend
        workbook.Save("DashboardChart_NoLegend.xlsx");
    }
}
