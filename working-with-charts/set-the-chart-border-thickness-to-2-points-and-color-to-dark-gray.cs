// Title: Aspose.Cells for .NET – Set Chart Border Thickness to 2 pt and Dark Gray Color (C#)
// Description: Creates a workbook, adds sample data, inserts a column chart, and customizes the ChartArea border by setting its WeightPt to 2.0 and Color to DarkGray before saving as ChartWithCustomBorder.xlsx.
// Keywords: Aspose.Cells | C# chart border | ChartArea.Border | WeightPt | Color.DarkGray | set chart border thickness | custom chart styling | Excel chart formatting .NET
// Common Searches: Aspose.Cells change chart border color C# | set chart border weight points Aspose.Cells | how to add dark gray border to Excel chart using Aspose | C# example chart area border styling Aspose.Cells | modify chart area line properties .NET
// Developer Intent: Add a 2‑point dark gray border to a chart’s chart area.
// Use Cases: Generate a column chart with a prominent dark gray border for clearer visual separation in automated reports. | Apply a consistent 2‑point border style to all charts in a workbook to meet corporate branding standards. | Export Excel charts with custom borders that match the design language of a web dashboard.
// AI Prompts: Show me C# code to set a chart's border thickness to 2 points and color to dark gray using Aspose.Cells. | Provide an Aspose.Cells example that modifies ChartArea.Border.WeightPt and Color in .NET. | Explain how to customize chart area borders (weight, color) for different chart types with Aspose.Cells for C#.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsChartBorderDemo
{
    // Creates a workbook, adds sample data, inserts a column chart, and customizes the ChartArea border by setting its WeightPt to 2.0 and Color to DarkGray before saving as ChartWithCustomBorder.xlsx.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate some sample data for the chart
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["A2"].PutValue("A");
            worksheet.Cells["A3"].PutValue("B");
            worksheet.Cells["A4"].PutValue("C");
            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["B2"].PutValue(10);
            worksheet.Cells["B3"].PutValue(20);
            worksheet.Cells["B4"].PutValue(30);

            // Add a column chart to the worksheet
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = worksheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Set the chart border thickness to 2 points and color to dark gray
            // The border is accessed via the ChartArea's Border (Line) object
            Line chartBorder = chart.ChartArea.Border;
            chartBorder.WeightPt = 2.0;               // Thickness in points
            chartBorder.Color = Color.DarkGray;       // Border color

            // Save the workbook to a file
            workbook.Save("ChartWithCustomBorder.xlsx");
        }
    }
}
