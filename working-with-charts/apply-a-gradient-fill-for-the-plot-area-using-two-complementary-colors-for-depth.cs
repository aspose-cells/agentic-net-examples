// Title: Apply a Horizontal Two‑Color Gradient to a Chart Plot Area with Aspose.Cells for .NET
// Description: This C# example creates a workbook, adds a column chart, and uses the PlotArea.FillFormat to set a horizontal gradient that transitions from teal to orange via SetTwoColorGradient, then saves the file as an XLSX document.
// Keywords: Aspose.Cells | .NET chart styling | plot area gradient fill | two‑color gradient | horizontal gradient | SetTwoColorGradient | C# Excel chart example | Teal orange gradient | Excel visual design | chart background color
// Common Searches: Aspose.Cells set chart plot area gradient | C# horizontal gradient fill for Excel chart | How to use SetTwoColorGradient in Aspose.Cells | gradient background for chart area .NET | apply two‑color gradient to chart plot area
// Developer Intent: Add a horizontal two‑color gradient to a chart’s plot area using Aspose.Cells for .NET.
// Use Cases: Enhance the visual depth of column charts in automated reports. | Match corporate color schemes by applying complementary gradient backgrounds. | Create presentation‑ready dashboards with stylized chart areas without manual Excel editing.
// AI Prompts: Generate C# code that applies a vertical three‑color gradient to a chart plot area with Aspose.Cells. | Show how to change the gradient style and variant of an existing chart plot area after it has been created. | Explain how to define custom RGB colors for a two‑color gradient instead of using predefined Color values.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// This C# example creates a workbook, adds a column chart, and uses the PlotArea.FillFormat to set a horizontal gradient that transitions from teal to orange via SetTwoColorGradient, then saves the file as an XLSX document.
class PlotAreaGradientDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["A4"].PutValue("C");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["B4"].PutValue(30);

        // Add a column chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
        Chart chart = sheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", true);          // Values
        chart.NSeries.CategoryData = "A2:A4";      // Categories

        // Access the plot area fill format
        FillFormat plotAreaFill = chart.PlotArea.Area.FillFormat;

        // Ensure the fill type is set to gradient
        plotAreaFill.FillType = FillType.Gradient;

        // Apply a two‑color gradient using complementary colors (Teal and Orange)
        // GradientStyleType.Horizontal creates a left‑to‑right transition
        // Variant = 1 selects the first preset variant
        plotAreaFill.SetTwoColorGradient(
            Color.Teal,          // First color
            Color.Orange,        // Second color
            GradientStyleType.Horizontal,
            1);

        // Save the workbook with the gradient‑filled plot area
        workbook.Save("PlotAreaGradientDemo.xlsx");
    }
}
