// Title: Show major gridlines on the secondary value axis of a column chart using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that builds a column chart, assigns the second series to a secondary axis, and turns on its gridlines with Aspose.Cells. | Generate an Aspose.Cells snippet that sets a blue color for the gridlines on the secondary value axis in a column chart. | Provide a complete Aspose.Cells program that creates a workbook, adds two data series, plots one on the secondary axis, and enables gridlines for better comparison.
// Common Searches: asp.net aspose.cells show gridlines for secondary side in column chart | c# show gridlines for the secondary side of a chart with Aspose.Cells | adjust gridline color on secondary side of Aspose.Cells column chart | aspose.cells .net guide to adding gridlines to a chart's secondary side
// Tags: Aspose.Cells secondary axis major gridlines C# | column chart secondary value axis formatting Aspose.Cells | activate gridlines for secondary value axis Aspose.Cells .NET | change secondary axis gridline hue Aspose.Cells | Aspose.Cells column chart with dual axes

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, adds sample data, builds a column chart with two series, plots the second series on the secondary value axis, makes the secondary axis major gridlines visible and colors them blue, then saves the file as SecondaryAxisMajorGridlines.xlsx.
class ShowSecondaryAxisMajorGridlines
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["A2"].PutValue("A");
        worksheet.Cells["A3"].PutValue("B");
        worksheet.Cells["A4"].PutValue("C");

        worksheet.Cells["B1"].PutValue("Series 1");
        worksheet.Cells["B2"].PutValue(10);
        worksheet.Cells["B3"].PutValue(20);
        worksheet.Cells["B4"].PutValue(30);

        worksheet.Cells["C1"].PutValue("Series 2");
        worksheet.Cells["C2"].PutValue(100);
        worksheet.Cells["C3"].PutValue(200);
        worksheet.Cells["C4"].PutValue(300);

        // Add a column chart
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
        Chart chart = worksheet.Charts[chartIndex];

        // Add two series to the chart
        chart.NSeries.Add("B2:B4", true);   // First series
        chart.NSeries.Add("C2:C4", true);   // Second series
        chart.NSeries.CategoryData = "A2:A4";

        // Plot the second series on the secondary value axis
        chart.NSeries[1].PlotOnSecondAxis = true;

        // Show major gridlines on the secondary value axis
        chart.SecondValueAxis.MajorGridLines.IsVisible = true;
        // Optional: set gridline color for better visibility
        chart.SecondValueAxis.MajorGridLines.Color = Color.Blue;

        // Save the workbook
        workbook.Save("SecondaryAxisMajorGridlines.xlsx");
    }
}
