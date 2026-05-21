using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

class RadialGradientChart
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
        sheet.Cells["B3"].PutValue(30);
        sheet.Cells["B4"].PutValue(20);

        // Add a column chart
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 15);
        Chart chart = sheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Access the first series and set its fill type to gradient
        Series series = chart.NSeries[0];
        series.Area.FillFormat.FillType = FillType.Gradient;

        // Configure the gradient as radial
        GradientFill gradientFill = series.Area.FillFormat.GradientFill;
        gradientFill.SetGradient(GradientFillType.Radial, 0, GradientDirectionType.FromCenter);

        // Clear any existing stops (optional, ensures a clean start)
        gradientFill.GradientStops.Clear();

        // Add three gradient stops for a smooth transition
        // Position is expressed as a percentage (0.0 to 1.0)
        // Alpha is 0-255 (255 = fully opaque)
        gradientFill.GradientStops.Add(0.0, Color.Red, 255);    // Start with red
        gradientFill.GradientStops.Add(0.5, Color.Green, 255);  // Middle green
        gradientFill.GradientStops.Add(1.0, Color.Blue, 255);   // End with blue

        // Save the workbook
        workbook.Save("RadialGradientChart.xlsx");
    }
}