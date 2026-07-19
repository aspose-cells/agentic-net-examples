// Title: Get Chart X/Y Position and Size with Aspose.Cells for .NET (C#)
// Description: This example creates a workbook, adds a column chart, calls chart.Calculate() to refresh layout, then reads the chart's ChartObject.X, ChartObject.Y, Width and Height properties (pixel units) and writes the values before saving the file.
// Keywords: Aspose.Cells | ChartObject X | ChartObject Y | chart position pixels | C# | .NET | retrieve chart coordinates | chart size Aspose | chart alignment | Excel chart location
// Common Searches: Aspose.Cells get chart X coordinate | how to read chart Y position in C# | retrieve chart pixel location Aspose.Cells | chart object width and height .NET | C# Aspose.Cells chart alignment example
// Developer Intent: Obtain the pixel‑level X and Y coordinates (and dimensions) of a chart object for precise placement within a worksheet.
// Use Cases: Align multiple charts programmatically by comparing their X/Y offsets. | Maintain exact layout when exporting a worksheet to PDF or image. | Adjust chart position dynamically after data changes or user interaction.
// AI Prompts: Generate C# code that moves an Aspose.Cells chart to a new X,Y location after reading its current coordinates. | Show how to loop through all charts in a workbook and log each chart's X, Y, width, and height. | Explain converting Aspose.Cells chart pixel coordinates to points for high‑resolution printing.

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;

// This example creates a workbook, adds a column chart, calls chart.Calculate() to refresh layout, then reads the chart's ChartObject.X, ChartObject.Y, Width and Height properties (pixel units) and writes the values before saving the file.
class RetrieveChartCoordinates
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["A2"].PutValue("A");
        worksheet.Cells["A3"].PutValue("B");
        worksheet.Cells["A4"].PutValue("C");
        worksheet.Cells["B1"].PutValue("Value");
        worksheet.Cells["B2"].PutValue(10);
        worksheet.Cells["B3"].PutValue(20);
        worksheet.Cells["B4"].PutValue(30);

        // Add a column chart to the worksheet
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = worksheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Ensure the chart layout is calculated so position properties are valid
        chart.Calculate();

        // Retrieve the chart object's position (in pixels) relative to the worksheet
        int chartX = chart.ChartObject.X;          // Horizontal offset from the left border
        int chartY = chart.ChartObject.Y;          // Vertical offset from the top border

        // Optionally retrieve size information
        int chartWidth = chart.ChartObject.Width;  // Width in pixels
        int chartHeight = chart.ChartObject.Height;// Height in pixels

        // Output the coordinates and size
        Console.WriteLine($"Chart Position -> X: {chartX} px, Y: {chartY} px");
        Console.WriteLine($"Chart Size -> Width: {chartWidth} px, Height: {chartHeight} px");

        // Save the workbook
        workbook.Save("ChartCoordinates.xlsx");
    }
}
