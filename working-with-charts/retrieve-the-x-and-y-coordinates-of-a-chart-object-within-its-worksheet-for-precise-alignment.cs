// Title: Get chart X/Y pixel coordinates and anchor cell indices with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to read a chart's exact placement in an Excel worksheet using Aspose.Cells. The example shows setting ChartObject.X and ChartObject.Y, then retrieving those pixel offsets together with UpperLeftRow and UpperLeftColumn to determine the chart's position before saving the file.
// Keywords: Aspose.Cells chart position | ChartObject X Y C# | retrieve chart pixel coordinates | chart anchor row column Aspose | Excel chart placement .NET
// Common Searches: Aspose.Cells get chart pixel location | C# read chart X and Y coordinates | how to find chart anchor cell in Excel using Aspose | retrieve chart placement properties Aspose.Cells | chart object position properties .NET
// Developer Intent: Extract the pixel‑level X/Y offsets and the upper‑left cell reference of a chart for precise layout control.
// Use Cases: Programmatically align multiple charts by comparing their X/Y values. | Generate a layout report that lists each chart’s exact screen coordinates and anchor cell. | Adjust chart placement dynamically based on surrounding content size.
// AI Prompts: Create a C# function that returns X, Y, UpperLeftRow, and UpperLeftColumn for any Aspose.Cells chart. | Write code to move a chart to new pixel coordinates while keeping its original anchor cell unchanged. | Show how to position two charts side‑by‑side by calculating and setting their ChartObject.X values.

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;

// Demonstrates how to read a chart's exact placement in an Excel worksheet using Aspose.Cells. The example shows setting ChartObject.X and ChartObject.Y, then retrieving those pixel offsets together with UpperLeftRow and UpperLeftColumn to determine the chart's position before saving the file.
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

        // Optionally set a known position for the chart (pixels from worksheet edges)
        chart.ChartObject.X = 150; // Horizontal offset in pixels
        chart.ChartObject.Y = 80;  // Vertical offset in pixels

        // Retrieve the chart's X and Y coordinates (pixel units)
        int chartPosX = chart.ChartObject.X;
        int chartPosY = chart.ChartObject.Y;

        // Retrieve the anchor cell indices (row and column) if needed
        int anchorRow = chart.ChartObject.UpperLeftRow;
        int anchorColumn = chart.ChartObject.UpperLeftColumn;

        // Output the coordinates
        Console.WriteLine($"Chart position (pixels): X = {chartPosX}, Y = {chartPosY}");
        Console.WriteLine($"Chart anchored at cell: Row = {anchorRow}, Column = {anchorColumn}");

        // Save the workbook
        workbook.Save("ChartCoordinates.xlsx");
    }
}
