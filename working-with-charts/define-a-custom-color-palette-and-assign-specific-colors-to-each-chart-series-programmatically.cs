// Title: Aspose.Cells .NET: Define a Custom Color Palette and Assign Colors to Chart Series
// Description: Creates a workbook, adds sample data, replaces the first two palette entries with custom RGB colors, inserts a column chart, binds data, sets each series' foreground color, disables varied point colors, and saves the file as an Excel workbook.
// Keywords: Aspose.Cells custom chart colors | change palette Aspose.Cells .NET | assign series color chart | disable varied colors Aspose.Cells | column chart custom palette C# | Excel chart color palette programmatic
// Common Searches: how to change chart palette in Aspose.Cells | set specific RGB colors for chart series Aspose.Cells | prevent varied point colors in Aspose.Cells column chart | customize Excel chart colors with C# Aspose.Cells
// Developer Intent: Generate an Excel chart with a bespoke color palette and fixed series colors using Aspose.Cells for .NET.
// Use Cases: Apply corporate brand colors to automatically generated charts. | Ensure consistent series colors across multiple charts in a reporting pipeline. | Create charts where each series retains a single color regardless of data point variations.
// AI Prompts: Show me C# code that modifies the workbook palette and assigns custom RGB colors to each chart series in Aspose.Cells. | Provide an example of creating a column chart with a custom color palette, explicit series colors, and IsColorVaried disabled using Aspose.Cells for .NET. | Explain how palette indices correspond to chart series colors and how to keep series colors uniform in Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// Creates a workbook, adds sample data, replaces the first two palette entries with custom RGB colors, inserts a column chart, binds data, sets each series' foreground color, disables varied point colors, and saves the file as an Excel workbook.
class CustomPaletteChart
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet ws = workbook.Worksheets[0];

        // Populate sample data for the chart
        ws.Cells["A1"].PutValue("Category");
        ws.Cells["A2"].PutValue("Jan");
        ws.Cells["A3"].PutValue("Feb");
        ws.Cells["A4"].PutValue("Mar");

        ws.Cells["B1"].PutValue("Series1");
        ws.Cells["B2"].PutValue(10);
        ws.Cells["B3"].PutValue(20);
        ws.Cells["B4"].PutValue(30);

        ws.Cells["C1"].PutValue("Series2");
        ws.Cells["C2"].PutValue(15);
        ws.Cells["C3"].PutValue(25);
        ws.Cells["C4"].PutValue(35);

        // Define custom colors and replace the default palette entries (indices 0 and 1)
        // These indices correspond to the first two colors used by chart series
        workbook.ChangePalette(Color.FromArgb(79, 129, 189), 0); // Custom blue
        workbook.ChangePalette(Color.FromArgb(192, 80, 77), 1); // Custom red

        // Add a column chart to the worksheet
        int chartIdx = ws.Charts.Add(ChartType.Column, 6, 0, 20, 10);
        Chart chart = ws.Charts[chartIdx];

        // Set the data range for the series and categories
        chart.NSeries.Add("B1:C4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Assign the custom colors to each series explicitly
        chart.NSeries[0].Area.ForegroundColor = Color.FromArgb(79, 129, 189); // Series1 color
        chart.NSeries[1].Area.ForegroundColor = Color.FromArgb(192, 80, 77); // Series2 color

        // Ensure that the series colors are not overridden by varied point colors
        chart.NSeries.IsColorVaried = false;

        // Save the workbook with the customized chart
        workbook.Save("CustomPaletteChart.xlsx");
    }
}
