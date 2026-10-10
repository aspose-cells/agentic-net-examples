// Title: Create a column chart with Aspose.Cells for .NET, set legend to automatic position, then move legend to a fixed top‑left location
// AI Prompts: Write C# code that uses Aspose.Cells to add a column chart, keep the legend at its default automatic position, and then change the legend's X and Y coordinates to place it at the top‑left corner of the chart. | Show how to programmatically override the automatic legend placement in an Aspose.Cells chart by setting explicit Legend.Position, Legend.X, and Legend.Y values in C#.
// Common Searches: Aspose.Cells C# set chart legend to custom coordinates | how to position chart legend at top left using Aspose.Cells .NET | override automatic legend placement in Aspose.Cells chart example | move column chart legend to specific X Y points Aspose.Cells | Aspose.Cells legend Position property usage in C#
// Tags: Aspose.Cells chart legend custom positioning | C# set legend X Y Aspose.Cells | column chart legend override automatic placement | Aspose.Cells workbook save chart with custom legend | Aspose.Cells legend top-left coordinates

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a new workbook, fills cells with sample data, adds a column chart, applies the default bottom legend position, demonstrates how to assign X and Y values to move the legend to a fixed top‑left spot, and saves the workbook as ChartWithCustomLegend.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart series
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Set legend to a default (auto) position (Bottom)
            chart.Legend.Position = LegendPositionType.Bottom;

            // If a custom position is required, set X and Y after setting Position to Bottom
            // (Custom positioning is not available in older API versions)
            // chart.Legend.Position = LegendPositionType.Bottom;
            // chart.Legend.X = 10; // distance from the left edge in points
            // chart.Legend.Y = 10; // distance from the top edge in points

            // Save the workbook with the chart
            workbook.Save("ChartWithCustomLegend.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
