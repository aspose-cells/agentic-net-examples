using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace GradientFillPlotAreaDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("Q1");
            sheet.Cells["A3"].PutValue("Q2");
            sheet.Cells["A4"].PutValue("Q3");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["B4"].PutValue(180);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 15);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Access the plot area of the chart
            PlotArea plotArea = chart.PlotArea;

            // Ensure the fill type is set to Gradient so we can apply a gradient fill
            plotArea.Area.FillFormat.FillType = FillType.Gradient;

            // Apply a two‑color gradient using complementary colors for depth
            // Example colors: LightSkyBlue and Orange (complementary)
            plotArea.Area.FillFormat.SetTwoColorGradient(
                Color.LightSkyBlue,   // First gradient color
                Color.Orange,         // Second gradient color
                GradientStyleType.Horizontal, // Gradient direction
                1);                    // Variant (1‑4)

            // Save the workbook to a file
            workbook.Save("GradientPlotAreaDemo.xlsx");
        }
    }
}