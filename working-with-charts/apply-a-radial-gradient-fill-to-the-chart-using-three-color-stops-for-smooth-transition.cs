using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsGradientDemo
{
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

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Access the first series
            Series series = chart.NSeries[0];

            // Set the fill type to gradient to enable gradient properties
            series.Area.FillFormat.FillType = FillType.Gradient;

            // Configure the gradient as radial
            GradientFill gradientFill = series.Area.FillFormat.GradientFill;
            gradientFill.SetGradient(GradientFillType.Radial, 0, GradientDirectionType.FromCenter);

            // Clear any existing gradient stops
            gradientFill.GradientStops.Clear();

            // Add three color stops for a smooth transition
            // Position values are percentages (0.0 to 1.0)
            gradientFill.GradientStops.Add(0.0, Color.Red, 255);    // Opaque red at start
            gradientFill.GradientStops.Add(0.5, Color.Yellow, 255); // Opaque yellow at middle
            gradientFill.GradientStops.Add(1.0, Color.Green, 255);  // Opaque green at end

            // Save the workbook
            workbook.Save("RadialGradientChart.xlsx");
        }
    }
}