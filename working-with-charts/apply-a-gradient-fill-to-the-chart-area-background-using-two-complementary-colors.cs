using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsChartAreaGradient
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

            // Access the fill format of the chart's plot area (the background of the chart area)
            FillFormat plotAreaFill = chart.PlotArea.Area.FillFormat;

            // Ensure the fill type is set to gradient
            plotAreaFill.FillType = FillType.Gradient;

            // Apply a two‑color gradient using complementary colors (Blue and Orange)
            // GradientStyleType.Horizontal creates a left‑to‑right transition
            // Variant = 1 selects the first variant of the gradient style
            plotAreaFill.SetTwoColorGradient(
                Color.Blue,          // First color
                Color.Orange,        // Second (complementary) color
                GradientStyleType.Horizontal,
                1);

            // Save the workbook with the gradient‑filled chart area
            workbook.Save("ChartAreaGradient.xlsx");
        }
    }
}