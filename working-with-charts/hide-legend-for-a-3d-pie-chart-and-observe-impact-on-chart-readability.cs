// Title: Hide the legend of a 3‑D pie chart in an Excel workbook with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# hide legend on a 3D pie chart example | remove legend from Pie3D chart using Aspose.Cells .NET | C# code to disable chart legend for 3‑D pie chart in Excel with Aspose.Cells | how to turn off legend visibility for a 3D pie chart in Aspose.Cells | Aspose.Cells chart readability without legend for Pie3D type
// Tags: hide chart legend Aspose.Cells | Pie3D chart customization Aspose.Cells | Excel legend visibility control C# | Aspose.Cells 3D pie chart formatting | programmatic chart legend removal .NET

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;

// The sample creates a new workbook, adds sample data, inserts a 3‑D pie chart (ChartType.Pie3D), sets ShowLegend to false to hide the legend, and saves the file as 3DPieChart_NoLegend.xlsx, demonstrating how legend visibility affects chart readability.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Fill sample data for the 3‑D pie chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B4"].PutValue(50);

            // Add a 3‑D pie chart (type Pie3D) to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Pie3D, 5, 0, 20, 5);
            Chart chart = sheet.Charts[chartIndex];

            // Define the data series (values) for the chart
            chart.NSeries.Add("B2:B4", true);
            // Category data line removed because the Series class in the current Aspose.Cells version
            // does not expose a CategoryData property. The chart will use default categories.

            // Hide the legend to observe its impact on readability
            chart.ShowLegend = false;

            // Save the workbook to a file
            workbook.Save("3DPieChart_NoLegend.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
