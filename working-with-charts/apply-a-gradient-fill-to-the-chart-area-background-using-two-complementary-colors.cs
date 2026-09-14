// Title: How to apply a horizontal two‑color gradient fill to a chart’s plot area using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that creates an Excel workbook, adds a column chart, and sets the chart’s plot area background to a horizontal gradient from blue to orange with Aspose.Cells. | Write a C# example that demonstrates configuring a two‑color gradient fill (horizontal) for the plot area of a chart in Aspose.Cells. | Provide a step‑by‑step C# snippet to apply a complementary color gradient to a chart area in an Excel file using Aspose.Cells.
// Common Searches: Aspose.Cells C# set horizontal gradient fill for chart plot area | How to use SetTwoColorGradient for Excel chart background in .NET | C# example applying blue to orange gradient to chart area with Aspose.Cells | Create column chart with gradient background using Aspose.Cells for .NET
// Tags: Aspose.Cells chart plot area gradient | C# SetTwoColorGradient usage | horizontal two‑color gradient Excel chart | chart background gradient Aspose.Cells .NET | gradient style type horizontal Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// Demonstrates creating a workbook, adding sample data, inserting a column chart, and applying a horizontal two‑color gradient (blue to orange) to the chart’s plot area using Aspose.Cells for .NET, then saving the file as ChartAreaGradient.xlsx.
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

        // Access the plot area fill format
        FillFormat plotAreaFill = chart.PlotArea.Area.FillFormat;

        // Set fill type to gradient to enable gradient operations
        plotAreaFill.FillType = FillType.Gradient;

        // Apply a two‑color gradient using complementary colors (Blue & Orange)
        // GradientStyleType.Horizontal creates a left‑to‑right transition
        // Variant = 1 (first variant)
        plotAreaFill.SetTwoColorGradient(Color.Blue, Color.Orange, GradientStyleType.Horizontal, 1);

        // Save the workbook with the gradient‑filled chart area
        workbook.Save("ChartAreaGradient.xlsx");
    }
}
