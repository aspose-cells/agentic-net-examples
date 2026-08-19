// Title: Aspose.Cells C# – Set chart legend to top‑right corner and read its pixel coordinates
// Description: Shows how to create a workbook, add a column chart, position the legend at the plot‑area corner (top‑right) with LegendPositionType.Corner, refresh the layout using chart.Calculate, and retrieve the legend's XPixel and YPixel values before saving the file.
// Keywords: Aspose.Cells | C# chart legend position | LegendPositionType.Corner | retrieve legend XPixel | retrieve legend YPixel | column chart Aspose.Cells | Excel automation legend coordinates | chart.Calculate legend update
// Common Searches: Aspose.Cells set legend corner | Get legend pixel position C# Aspose.Cells | Chart legend top right Aspose.Cells example | How to read legend XPixel YPixel after chart.Calculate | Aspose.Cells legend placement API
// Developer Intent: Place a chart legend in the top‑right corner and obtain its pixel location programmatically.
// Use Cases: Guarantee consistent legend placement in automatically generated Excel reports. | Perform visual‑regression tests by comparing legend coordinates across versions. | Calculate offsets for other drawing objects to avoid overlapping the legend.
// AI Prompts: Provide C# code that sets a chart legend to the Corner position and returns its XPixel and YPixel values using Aspose.Cells. | Explain how to verify legend coordinates after calling chart.Calculate in Aspose.Cells for .NET. | Show an example of adjusting other shapes based on the legend's pixel coordinates in an Aspose.Cells workbook.

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;

// Shows how to create a workbook, add a column chart, position the legend at the plot‑area corner (top‑right) with LegendPositionType.Corner, refresh the layout using chart.Calculate, and retrieve the legend's XPixel and YPixel values before saving the file.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
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

        // Add a column chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = sheet.Charts[chartIndex];
        chart.SetChartDataRange("A1:B4", true);

        // Set the legend position to the top‑right corner of the plot area
        chart.Legend.Position = LegendPositionType.Corner;

        // Calculate the chart layout so that position properties are up‑to‑date
        chart.Calculate();

        // Retrieve the legend coordinates in pixels (XPixel, YPixel)
        int legendX = chart.Legend.XPixel;
        int legendY = chart.Legend.YPixel;

        // Verify by outputting the coordinates
        Console.WriteLine($"Legend positioned at Corner. Coordinates (pixels): X = {legendX}, Y = {legendY}");

        // Save the workbook
        workbook.Save("LegendTopRightCorner.xlsx");
    }
}
