// Title: Aspose.Cells for .NET – Apply a Two‑Color Gradient to a Chart Plot Area (C#)
// Description: C# code that builds a workbook, fills it with sample data, creates a column chart, and uses PlotArea.FillFormat.SetTwoColorGradient to paint a horizontal blue‑to‑orange gradient on the chart’s plot area. The file is saved as ChartAreaGradient.xlsx, showing how to style chart backgrounds with Aspose.Cells.
// Keywords: Aspose.Cells chart gradient fill | SetTwoColorGradient C# | chart plot area background Aspose.Cells | horizontal two‑color gradient .NET | C# Aspose.Cells chart styling | gradient style chart area | Aspose.Cells FillFormat example
// Common Searches: Aspose.Cells how to add gradient to chart plot area C# | SetTwoColorGradient example for column chart Aspose.Cells | C# code for horizontal gradient background in Excel chart | apply complementary colors gradient to Excel chart using Aspose | Aspose.Cells chart area fill format tutorial
// Developer Intent: Add a horizontal blue‑to‑orange two‑color gradient to the plot area of a column chart using Aspose.Cells for .NET.
// Use Cases: Improve visual impact of Excel dashboards by programmatically applying gradient fills to chart backgrounds. | Maintain a consistent color scheme across multiple charts in automated report generation. | Highlight data sections with complementary color gradients for clearer presentation in financial or sales reports.
// AI Prompts: Generate C# code with Aspose.Cells to set a vertical two‑color gradient from green to red on a chart's plot area. | Show how to customize gradient direction, style, and variant for a chart area background using Aspose.Cells in .NET.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// C# code that builds a workbook, fills it with sample data, creates a column chart, and uses PlotArea.FillFormat.SetTwoColorGradient to paint a horizontal blue‑to‑orange gradient on the chart’s plot area. The file is saved as ChartAreaGradient.xlsx, showing how to style chart backgrounds with Aspose.Cells.
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
        chart.NSeries.Add("B2:B4", true);          // values
        chart.NSeries.CategoryData = "A2:A4";      // categories

        // Apply a two‑color gradient to the chart area (plot area) background
        // Using complementary colors: Blue and Orange
        FillFormat chartAreaFill = chart.PlotArea.Area.FillFormat;
        chartAreaFill.SetTwoColorGradient(
            Color.Blue,          // first color
            Color.Orange,        // second (complementary) color
            GradientStyleType.Horizontal, // gradient direction
            1);                  // variant (1‑4)

        // Save the workbook
        workbook.Save("ChartAreaGradient.xlsx");
    }
}
