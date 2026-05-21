using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

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
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = sheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", true);          // Values
        chart.NSeries.CategoryData = "A2:A4";      // Categories

        // Access the fill format of the chart's plot area background
        FillFormat fillFormat = chart.PlotArea.Area.FillFormat;

        // Set the fill type to gradient (required before applying gradient settings)
        fillFormat.FillType = FillType.Gradient;

        // Apply a two‑color gradient using complementary colors (e.g., Orange and Blue)
        fillFormat.SetTwoColorGradient(
            Color.Orange,          // First color
            Color.Blue,            // Second color
            GradientStyleType.Horizontal, // Gradient direction
            1);                    // Variant (1‑4)

        // Save the workbook with the gradient‑filled chart area
        workbook.Save("ChartAreaGradient.xlsx");
    }
}