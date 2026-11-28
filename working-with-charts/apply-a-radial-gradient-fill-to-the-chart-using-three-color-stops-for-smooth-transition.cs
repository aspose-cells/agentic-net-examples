// Title: How to apply a radial two‑color gradient fill to a chart’s plot area using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells set radial gradient on chart plot area | Create column chart with two‑color gradient background using Aspose.Cells | Use SetTwoColorGradient and GradientStyleType.FromCenter in Aspose.Cells .NET | Apply radial gradient fill to Excel chart with Aspose.Cells API | Aspose.Cells example for chart background gradient styling
// Tags: Aspose.Cells FillFormat radial gradient | SetTwoColorGradient method C# Aspose.Cells | GradientStyleType.FromCenter example Aspose.Cells | column chart gradient fill Aspose.Cells | save workbook with chart gradient Aspose.Cells .NET

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace RadialGradientChartDemo
{
    // This C# example creates a workbook, adds sample data, inserts a column chart, and configures the chart’s plot area to a radial two‑color gradient (red to blue) by setting FillFormat.FillType to Gradient, calling SetTwoColorGradient with GradientStyleType.FromCenter, ensuring a solid fill pattern, and saving the workbook as an XLSX file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["B1"].PutValue("Value");
                worksheet.Cells["A2"].PutValue("Alpha");
                worksheet.Cells["B2"].PutValue(30);
                worksheet.Cells["A3"].PutValue("Beta");
                worksheet.Cells["B3"].PutValue(55);
                worksheet.Cells["A4"].PutValue("Gamma");
                worksheet.Cells["B4"].PutValue(80);

                // Add a column chart to the worksheet
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
                Chart chart = worksheet.Charts[chartIndex];

                // Set the data range for the chart
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Apply a radial gradient to the chart plot area
                // 1. Set the fill type to Gradient
                chart.PlotArea.Area.FillFormat.FillType = FillType.Gradient;

                // 2. Define a two‑color gradient (Aspose.Cells supports only two colors)
                //    Use GradientStyleType.FromCenter for a radial effect
                chart.PlotArea.Area.FillFormat.SetTwoColorGradient(
                    Color.Red,          // Inner color
                    Color.Blue,        // Outer color
                    GradientStyleType.FromCenter,
                    1);                 // Variant (valid for FromCenter)

                // Ensure the fill pattern is solid (no pattern overrides the gradient)
                chart.PlotArea.Area.FillFormat.Pattern = FillPattern.Solid;

                // Save the workbook to an XLSX file
                string outputPath = "RadialGradientChart.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
