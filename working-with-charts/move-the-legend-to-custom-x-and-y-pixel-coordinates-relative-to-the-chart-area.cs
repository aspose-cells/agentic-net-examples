// Title: Set a chart legend to exact X and Y pixel coordinates in Aspose.Cells for .NET (C#)
// AI Prompts: Place the legend of a column chart at X=150 px and Y=50 px relative to the chart area using Aspose.Cells in C#. | Configure a chart legend to a free (NotDocked) position and assign XPixel and YPixel values for precise placement with Aspose.Cells. | Programmatically move an Excel chart legend to custom pixel offsets in a .NET workbook using Aspose.Cells.
// Common Searches: asp.net aspose.cells set chart legend pixel position | c# Aspose.Cells legend XPixel YPixel example | how to move Excel chart legend to custom coordinates with Aspose.Cells | Aspose.Cells chart legend NotDocked placement tutorial
// Tags: set legend pixel position Aspose.Cells | chart legend NotDocked C# | custom legend coordinates Excel Aspose.Cells | column chart legend offset pixels

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook, adds sample data, inserts a column chart, sets the legend to NotDocked, moves it to XPixel = 150 and YPixel = 50, and saves the file as CustomLegendPosition.xlsx.
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
        chart.NSeries.Add("B2:B4", true);          // Values
        chart.NSeries.CategoryData = "A2:A4";      // Categories

        // Set legend to a free (not docked) position so pixel coordinates are respected
        chart.Legend.Position = LegendPositionType.NotDocked;

        // Move the legend to custom pixel coordinates relative to the chart area
        chart.Legend.XPixel = 150; // Horizontal offset in pixels
        chart.Legend.YPixel = 50;  // Vertical offset in pixels

        // Save the workbook with the customized legend position
        workbook.Save("CustomLegendPosition.xlsx");
    }
}
