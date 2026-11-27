// Title: How to apply a vertical two‑color gradient fill to a chart plot area using Aspose.Cells for .NET (C#)
// Common Searches: c# aspnet set vertical two‑color gradient for Excel chart plot area with Aspose.Cells | example of FillFormat.SetTwoColorGradient on a chart using Aspose.Cells .NET | apply complementary color gradient background to a column chart in C# with Aspose.Cells | Aspose.Cells chart plot area gradient fill vertical direction C# sample code | how to change chart plot area background to a blue‑to‑orange gradient in Aspose.Cells
// Tags: set two‑color gradient fill chart plot area Aspose.Cells | vertical gradient background Excel chart C# | FillFormat SetTwoColorGradient Aspose.Cells | gradient style vertical Aspose.Cells API | gradient variant parameter Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace ChartGradientDemo
{
    // The example creates a workbook, adds sample data and a column chart, then uses the chart's PlotArea.FillFormat to apply a vertical two‑color gradient (blue to orange) via SetTwoColorGradient, sets the fill type to Gradient, and saves the file as an XLSX workbook.
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
                worksheet.Cells["A3"].PutValue("Beta");
                worksheet.Cells["A4"].PutValue("Gamma");
                worksheet.Cells["B2"].PutValue(30);
                worksheet.Cells["B3"].PutValue(55);
                worksheet.Cells["B4"].PutValue(80);

                // Add a column chart to the worksheet
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
                Chart chart = worksheet.Charts[chartIndex];

                // Set the data source for the chart
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Optional: give the chart a title
                chart.Title.Text = "Sample Column Chart";

                // Apply a two‑color gradient fill to the chart area background
                // Complementary colors: Blue and Orange
                FillFormat chartAreaFill = chart.PlotArea.Area.FillFormat;
                chartAreaFill.SetTwoColorGradient(
                    Color.Blue,          // First gradient color
                    Color.Orange,        // Second gradient color
                    GradientStyleType.Vertical, // Gradient direction
                    1);                  // Variant (1‑4)

                // Explicitly set the fill type to Gradient
                chartAreaFill.FillType = FillType.Gradient;

                // Save the workbook to a file
                string outputPath = "ChartGradientDemo.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);

                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
                Console.WriteLine("The chart area now has a vertical two‑color gradient (Blue → Orange).");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while creating the chart gradient:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
