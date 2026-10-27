// Title: How to set the outline of all chart series to the workbook theme’s Dark1 color using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells set chart series border to workbook theme Dark1 color | apply theme color to all series outlines in an Aspose.Cells chart | Aspose.Cells change column chart series outline using theme colors | set series border visibility and color programmatically with Aspose.Cells | retrieve workbook theme Dark1 color for chart series styling Aspose.Cells
// Tags: set chart series outline theme color Aspose.Cells | apply workbook theme Dark1 to chart borders C# | programmatic chart series border styling Aspose.Cells | retrieve theme color for chart series Aspose.Cells | chart series border visibility Aspose.Cells C#

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

// The example creates a new workbook, adds sample data and a column chart, then iterates through each series to set the series border color (e.g., to the workbook theme's Dark1 color) and makes the border visible before saving the file as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook wb = new Workbook();

            // Get the first worksheet
            Worksheet ws = wb.Worksheets[0];

            // Add sample data for the chart
            ws.Cells["A1"].PutValue("Category");
            ws.Cells["B1"].PutValue("Value");
            ws.Cells["A2"].PutValue("A");
            ws.Cells["A3"].PutValue("B");
            ws.Cells["A4"].PutValue("C");
            ws.Cells["B2"].PutValue(10);
            ws.Cells["B3"].PutValue(20);
            ws.Cells["B4"].PutValue(30);

            // Add a column chart
            int chartIndex = ws.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = ws.Charts[chartIndex];

            // Define the series and categories
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Use a default color (e.g., Black) for the series outline
            Color outlineColor = Color.Black;

            // Apply outline color to all chart series
            foreach (Series series in chart.NSeries)
            {
                series.Border.Color = outlineColor;
                series.Border.IsVisible = true;
            }

            // Save the workbook
            wb.Save("output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
