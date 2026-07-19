// Title: Aspose.Cells C# – Apply a Two‑Color Gradient to a Chart Area Background
// Description: Creates a workbook, adds sample data, inserts a column chart, sets the plot‑area fill type to Gradient, and applies a horizontal blue‑to‑orange two‑color gradient (variant 1) before saving as ChartAreaGradient.xlsx.
// Keywords: Aspose.Cells | C# | chart area gradient | two color gradient | SetTwoColorGradient | plot area fill | Excel chart styling | horizontal gradient | blue orange gradient | example code
// Common Searches: Aspose.Cells set two‑color gradient for chart background C# | how to apply horizontal gradient to chart area using Aspose.Cells | C# code for gradient fill of Excel chart plot area | Aspose.Cells chart area fill type gradient example
// Developer Intent: Add a horizontal two‑color gradient background to the plot area of an Excel chart using Aspose.Cells for .NET.
// Use Cases: Design sales dashboards where the chart background follows corporate blue‑orange branding. | Standardize visual style across multiple workbooks by programmatically applying gradient fills. | Generate automated reports with visually appealing charts that use complementary gradient colors.
// AI Prompts: Show how to change the gradient direction to vertical and choose custom colors for the chart area in Aspose.Cells C#. | Provide C# code to apply a three‑color gradient to a line chart’s plot area with Aspose.Cells. | Explain how to read, modify, or replace an existing gradient fill of a chart area after loading a workbook.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// Creates a workbook, adds sample data, inserts a column chart, sets the plot‑area fill type to Gradient, and applies a horizontal blue‑to‑orange two‑color gradient (variant 1) before saving as ChartAreaGradient.xlsx.
class GradientChartAreaDemo
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

        // Insert a column chart
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
        Chart chart = sheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", true);          // Values
        chart.NSeries.CategoryData = "A2:A4";      // Categories

        // Access the plot area (chart area) fill format
        FillFormat chartAreaFill = chart.PlotArea.Area.FillFormat;

        // Ensure the fill type is set to gradient
        chartAreaFill.FillType = FillType.Gradient;

        // Apply a two‑color gradient using complementary colors (Blue and Orange)
        chartAreaFill.SetTwoColorGradient(
            Color.Blue,          // First color
            Color.Orange,        // Second color
            GradientStyleType.Horizontal, // Gradient direction
            1);                  // Variant (1‑4)

        // Save the workbook with the gradient‑filled chart area
        workbook.Save("ChartAreaGradient.xlsx");
    }
}
