// Title: C# – Apply a White‑to‑Gray Vertical Gradient to a Chart Plot Area with Aspose.Cells
// Description: Creates a workbook, adds sample data, inserts a column chart, and uses the chart's PlotArea.FillFormat.SetTwoColorGradient method to paint a vertical gradient that fades from white at the top to gray at the bottom, then saves the file as an .xlsx document.
// Keywords: Aspose.Cells C# chart gradient | SetTwoColorGradient plot area | vertical gradient fill Aspose.Cells | chart background gradient .NET | white to gray gradient Aspose.Cells
// Common Searches: Aspose.Cells apply vertical gradient to chart plot area | SetTwoColorGradient example C# Aspose.Cells | how to change chart plot area background Aspose.Cells .NET | gradient fill for Excel chart using Aspose.Cells | white gray gradient chart Aspose.Cells
// Developer Intent: Add a vertical white‑to‑gray gradient fill to the plot area of an Excel chart programmatically.
// Use Cases: Enhance the visual style of automatically generated reports by applying a consistent gradient to all chart plot areas. | Match corporate branding guidelines that require a light‑to‑dark background on chart graphics. | Create a template workbook where each new chart inherits a predefined gradient background without manual formatting.
// AI Prompts: Generate C# code that uses Aspose.Cells to set a vertical two‑color gradient on a chart's plot area. | Show how to change the gradient colors and direction for a chart plot area with Aspose.Cells for .NET. | Explain the steps to apply a custom gradient (e.g., blue to green) to an Excel chart using Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// Creates a workbook, adds sample data, inserts a column chart, and uses the chart's PlotArea.FillFormat.SetTwoColorGradient method to paint a vertical gradient that fades from white at the top to gray at the bottom, then saves the file as an .xlsx document.
class GradientPlotAreaDemo
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
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
        Chart chart = sheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Apply a vertical two‑color gradient (white at top, gray at bottom) to the plot area
        FillFormat plotAreaFill = chart.PlotArea.Area.FillFormat;
        plotAreaFill.SetTwoColorGradient(Color.White, Color.Gray, GradientStyleType.Vertical, 1);

        // Save the workbook
        workbook.Save("GradientPlotAreaDemo.xlsx");
    }
}
