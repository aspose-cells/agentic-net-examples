// Title: Aspose.Cells .NET – Apply a vertical white‑to‑gray gradient to a chart plot area
// Description: This C# example creates a workbook, adds sample data, inserts a column chart, and uses Chart.PlotArea.Area.FillFormat.SetTwoColorGradient to paint the plot area with a vertical white‑to‑gray two‑color gradient before saving the file as XLSX.
// Keywords: Aspose.Cells | C# chart gradient fill | plot area background | SetTwoColorGradient | vertical two‑color gradient | Excel chart styling .NET | Aspose.Cells for .NET examples | gradient style chart | chart plot area fill format | Aspose.Cells chart background
// Common Searches: Aspose.Cells vertical gradient chart C# | How to set chart plot area gradient with Aspose.Cells | C# SetTwoColorGradient example | Apply gradient to Excel chart using Aspose.Cells | Chart background gradient Aspose.Cells .NET
// Developer Intent: Add a vertical white‑to‑gray two‑color gradient to a chart's plot area.
// Use Cases: Enhance the visual appeal of column charts in automated Excel reports. | Match chart backgrounds to corporate color schemes without manual editing. | Batch‑style multiple charts in a workbook by applying a consistent gradient.
// AI Prompts: Generate C# code that uses Aspose.Cells to apply a horizontal blue‑to‑green gradient to a chart plot area. | Show how to loop through all charts in a workbook and set a light‑gray to dark‑gray vertical gradient on each plot area. | Explain the parameters of SetTwoColorGradient and demonstrate how to create radial and diagonal gradients with Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// This C# example creates a workbook, adds sample data, inserts a column chart, and uses Chart.PlotArea.Area.FillFormat.SetTwoColorGradient to paint the plot area with a vertical white‑to‑gray two‑color gradient before saving the file as XLSX.
class GradientPlotAreaDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add some sample data for the chart (required for a valid chart)
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["A4"].PutValue("C");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["B4"].PutValue(30);

        // Add a column chart
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data range for the chart
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Apply a vertical two‑color gradient (white at top, gray at bottom) to the plot area
        FillFormat plotAreaFill = chart.PlotArea.Area.FillFormat;
        plotAreaFill.SetTwoColorGradient(Color.White, Color.Gray, GradientStyleType.Vertical, 1);

        // Save the workbook
        workbook.Save("GradientPlotAreaDemo.xlsx");
    }
}
