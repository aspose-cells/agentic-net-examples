// Title: Apply a radial gradient fill with three color stops to a column chart series in Aspose.Cells for .NET
// AI Prompts: Generate C# code that creates a column chart and sets its series area fill to a radial gradient with red, green, and blue stops using Aspose.Cells. | Write a snippet that configures a GradientFill object for a chart series to use GradientFillType.Radial and adds three opacity‑100 color stops at positions 0, 0.5, and 1. | Produce an example that saves an Excel workbook containing a column chart whose series area displays a smooth radial transition from red to green to blue. | Create a C# program that clears existing gradient stops and adds custom stops to achieve a three‑color radial gradient on a chart series with Aspose.Cells.
// Common Searches: Aspose.Cells C# set radial gradient fill on chart series with multiple color stops | how to add three gradient stops to a column chart using Aspose.Cells .NET | example of radial gradient fill type for Excel chart series in C# | configure chart series area gradient direction from center Aspose.Cells | apply custom color gradient to Excel chart using Aspose.Cells API
// Tags: radial gradient fill chart series Aspose.Cells | add multiple gradient stops to chart area C# | column chart series gradient fill Aspose.Cells .NET | GradientFillType.Radial usage Aspose.Cells | custom three‑color gradient Excel chart C#

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// The example creates a workbook, adds sample data, inserts a column chart, and applies a radial gradient fill to the chart series area using three opaque color stops (red, green, blue). The workbook is saved as RadialGradientChart.xlsx.
class RadialGradientChartExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("Category 1");
        sheet.Cells["A3"].PutValue("Category 2");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(20);

        // Add a column chart
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
        Chart chart = sheet.Charts[chartIndex];
        chart.SetChartDataRange("A1:B3", true);

        // Access the first series of the chart
        Series series = chart.NSeries[0];

        // Set the fill type to gradient to enable gradient properties
        series.Area.FillFormat.FillType = FillType.Gradient;

        // Obtain the GradientFill object
        GradientFill gradientFill = series.Area.FillFormat.GradientFill;

        // Configure the gradient as radial
        gradientFill.SetGradient(GradientFillType.Radial, 0, GradientDirectionType.FromCenter);

        // Clear any existing stops (optional, ensures a clean state)
        gradientFill.GradientStops.Clear();

        // Add three color stops for a smooth transition
        gradientFill.GradientStops.Add(0.0, Color.Red, 255);    // Start with opaque red
        gradientFill.GradientStops.Add(0.5, Color.Green, 255);  // Middle with opaque green
        gradientFill.GradientStops.Add(1.0, Color.Blue, 255);   // End with opaque blue

        // Save the workbook with the chart
        workbook.Save("RadialGradientChart.xlsx");
    }
}
