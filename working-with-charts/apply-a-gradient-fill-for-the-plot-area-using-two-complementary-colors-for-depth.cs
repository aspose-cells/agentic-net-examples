// Title: How to apply a vertical two‑color gradient (blue to orange) to a chart PlotArea using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set vertical two‑color gradient on chart PlotArea | Create Excel column chart with gradient background in PlotArea using Aspose.Cells | Set gradient fill for chart PlotArea in .NET Aspose.Cells example | Apply blue to orange gradient to Excel chart area with Aspose.Cells API | How to use SetTwoColorGradient for PlotArea in Aspose.Cells C#
// Tags: set two‑color gradient plotarea Aspose.Cells | gradient style vertical Aspose.Cells | column chart plotarea formatting C# | SetTwoColorGradient API usage | apply blue orange gradient Excel chart Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace GradientPlotAreaDemo
{
    // The sample creates a workbook, adds sample data, inserts a column chart, accesses the chart's PlotArea, and uses the SetTwoColorGradient method with a vertical blue‑to‑orange gradient before saving the file as an XLSX workbook.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["A2"].PutValue("Alpha");
                sheet.Cells["A3"].PutValue("Beta");
                sheet.Cells["A4"].PutValue("Gamma");
                sheet.Cells["B2"].PutValue(30);
                sheet.Cells["B3"].PutValue(55);
                sheet.Cells["B4"].PutValue(80);

                // Add a column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 15);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data range for the chart
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Access the PlotArea of the chart
                PlotArea plotArea = chart.PlotArea;
                Area area = plotArea.Area;

                // Apply a two‑color vertical gradient (Blue → Orange) to the plot area
                area.FillFormat.SetTwoColorGradient(
                    Color.Blue,          // First color
                    Color.Orange,        // Second color
                    GradientStyleType.Vertical,
                    1);

                // Save the workbook
                string outputPath = "GradientPlotArea.xlsx";
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
