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

            // Populate some sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Q1");
            sheet.Cells["A3"].PutValue("Q2");
            sheet.Cells["A4"].PutValue("Q3");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["B4"].PutValue(180);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Access the plot area of the chart
            PlotArea plotArea = chart.PlotArea;

            // Get the FillFormat of the plot area
            FillFormat plotAreaFill = plotArea.Area.FillFormat;

            // Apply a two‑color gradient using complementary colors (Blue and Orange)
            // GradientStyleType.Horizontal creates a left‑to‑right transition
            // Variant = 1 selects the first preset variant
            plotAreaFill.SetTwoColorGradient(
                Color.Blue,          // First color
                Color.Orange,        // Second (complementary) color
                GradientStyleType.Horizontal,
                1);

            // Save the workbook to a file
            workbook.Save("PlotAreaGradientDemo.xlsx");
        }
    }
}