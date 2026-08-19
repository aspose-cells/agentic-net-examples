// Title: Aspose.Cells .NET – Apply a Radial Gradient Fill with Three Color Stops to a Chart Series
// Description: C# example that creates a workbook, adds a column chart, and configures the first series to use a radial gradient fill with red, green, and blue stops (0%, 50%, 100%). The gradient type is set to Radial and the file is saved as RadialGradientChart.xlsx.
// Keywords: Aspose.Cells | .NET | C# | radial gradient fill | chart series gradient | gradient stops | Excel chart styling | GradientFillType.Radial | GradientDirectionType.FromCenter | example code
// Common Searches: Aspose.Cells radial gradient chart series C# | how to add multiple gradient stops to Excel chart using Aspose.Cells | set radial gradient fill for column chart in .NET | gradient fill example Aspose.Cells chart series
// Developer Intent: Add a radial gradient with three distinct color stops to a chart series in an Excel workbook using Aspose.Cells for .NET.
// Use Cases: Design eye‑catching column charts where data series transition smoothly between colors. | Generate Excel reports with custom gradient styling for brand‑consistent visuals. | Create presentation‑ready charts that differentiate categories through radial color blends.
// AI Prompts: Show how to make the gradient stops semi‑transparent while keeping the radial effect. | Provide code to apply a radial gradient with custom colors to multiple series in the same chart. | Explain how to modify the radius and direction of a radial gradient on a chart series using Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsGradientDemo
{
    // C# example that creates a workbook, adds a column chart, and configures the first series to use a radial gradient fill with red, green, and blue stops (0%, 50%, 100%). The gradient type is set to Radial and the file is saved as RadialGradientChart.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Configure the first series to use a radial gradient with three color stops
            Series series = chart.NSeries[0];
            series.Area.FillFormat.FillType = FillType.Gradient;                     // Enable gradient fill
            GradientFill gradientFill = series.Area.FillFormat.GradientFill;         // Get GradientFill object

            // Clear any existing gradient stops
            gradientFill.GradientStops.Clear();

            // Add three gradient stops: Red at start, Green at middle, Blue at end
            gradientFill.GradientStops.Add(0.0, Color.Red, 255);   // Position 0% (opaque red)
            gradientFill.GradientStops.Add(0.5, Color.Green, 255); // Position 50% (opaque green)
            gradientFill.GradientStops.Add(1.0, Color.Blue, 255);  // Position 100% (opaque blue)

            // Set the gradient type to radial
            gradientFill.SetGradient(GradientFillType.Radial, 0, GradientDirectionType.FromCenter);

            // Save the workbook
            workbook.Save("RadialGradientChart.xlsx");
        }
    }
}
