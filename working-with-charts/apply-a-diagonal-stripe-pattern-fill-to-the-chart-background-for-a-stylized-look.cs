// Title: Add a diagonal stripe pattern fill to a chart area with Aspose.Cells for .NET (C#)
// Description: This example creates a workbook, inserts sample data, builds a column chart, and styles the chart's background (ChartArea) using a diagonal stripe pattern (FillPattern.LightDownwardDiagonal). It also shows how to set foreground and background colors before saving the file as an XLSX workbook.
// Keywords: Aspose.Cells chart pattern fill | C# diagonal stripe chart background | FillPattern.LightDownwardDiagonal | ChartArea FillType.Pattern | Aspose.Cells example chart styling | Excel chart background pattern .NET | pattern fill chart area Aspose
// Common Searches: Aspose.Cells set diagonal pattern for chart area | C# chart background pattern fill Aspose.Cells | How to apply LightDownwardDiagonal fill to Excel chart | Aspose.Cells chart area FillType.Pattern example | Add stripe pattern to chart background in .NET
// Developer Intent: Apply a diagonal stripe pattern to the chart area for visual styling.
// Use Cases: Enhance a column chart with a branded diagonal stripe background. | Create Excel reports where chart backgrounds need patterned fills for emphasis. | Demonstrate pattern‑based styling of chart areas in Aspose.Cells tutorials.
// AI Prompts: Show a C# code snippet that sets FillType.Pattern and FillPattern.LightDownwardDiagonal on a chart area using Aspose.Cells. | Explain how to change the foreground and background colors of a diagonal stripe pattern in an Aspose.Cells chart. | Provide alternatives for other pattern fills (e.g., LightUpwardDiagonal, DarkHorizontal) in Aspose.Cells chart styling.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsChartPatternDemo
{
    // This example creates a workbook, inserts sample data, builds a column chart, and styles the chart's background (ChartArea) using a diagonal stripe pattern (FillPattern.LightDownwardDiagonal). It also shows how to set foreground and background colors before saving the file as an XLSX workbook.
    public class Program
    {
        public static void Main()
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
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Apply a diagonal stripe pattern fill to the chart background (ChartArea)
            // Set the fill type to Pattern
            chart.ChartArea.Area.FillFormat.FillType = FillType.Pattern;
            // Choose a diagonal pattern, e.g., LightDownwardDiagonal
            chart.ChartArea.Area.FillFormat.PatternFill.Pattern = FillPattern.LightDownwardDiagonal;
            // Optional: set foreground and background colors for the pattern
            chart.ChartArea.Area.FillFormat.PatternFill.ForegroundColor = Color.LightBlue;
            chart.ChartArea.Area.FillFormat.PatternFill.BackgroundColor = Color.White;

            // Save the workbook
            workbook.Save("ChartWithDiagonalStripePattern.xlsx", SaveFormat.Xlsx);
        }
    }
}
