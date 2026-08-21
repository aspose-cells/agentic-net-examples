// Title: Set a White‑to‑Gray Vertical Gradient on a Chart Plot Area with Aspose.Cells for .NET
// Description: Shows how to create a workbook, add sample data, insert a column chart, and use FillFormat.SetTwoColorGradient to apply a white‑top, gray‑bottom vertical gradient to the chart’s plot area before saving the file.
// Keywords: Aspose.Cells chart gradient | SetTwoColorGradient C# | vertical gradient plot area | chart background styling Aspose.Cells | Excel chart fill format .NET
// Common Searches: Aspose.Cells apply gradient to chart plot area C# | vertical two‑color gradient chart Aspose.Cells example | white to gray gradient Excel chart using Aspose.Cells | set chart background gradient Aspose.Cells .NET
// Developer Intent: Programmatically add a white‑to‑gray vertical gradient to an Excel chart’s plot area.
// Use Cases: Enhance the visual design of column charts in automated reports. | Apply consistent gradient branding to chart backgrounds across multiple worksheets. | Bulk‑format chart plot areas in existing workbooks to match corporate style guidelines.
// AI Prompts: Generate C# code with Aspose.Cells that sets a vertical two‑color gradient (white to gray) on a chart’s plot area. | Explain how FillFormat.SetTwoColorGradient works for different chart types in Aspose.Cells for .NET. | Provide examples of customizing gradient direction and colors for Excel chart backgrounds using Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// Shows how to create a workbook, add sample data, insert a column chart, and use FillFormat.SetTwoColorGradient to apply a white‑top, gray‑bottom vertical gradient to the chart’s plot area before saving the file.
class Program
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
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Apply a vertical gradient fill (white at top, gray at bottom) to the plot area
        FillFormat plotAreaFill = chart.PlotArea.Area.FillFormat;
        plotAreaFill.SetTwoColorGradient(Color.White, Color.Gray, GradientStyleType.Vertical, 1);

        // Save the workbook
        workbook.Save("ChartPlotAreaGradient.xlsx");
    }
}
