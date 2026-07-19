// Title: C# – Aspose.Cells – Apply a Two‑Color Horizontal Gradient to a Chart Plot Area
// Description: Demonstrates how to create a workbook, add a column chart, and use the PlotArea.FillFormat.SetTwoColorGradient method to paint the plot area with a LightSkyBlue‑to‑DarkBlue horizontal gradient before saving the file as XLSX.
// Keywords: Aspose.Cells C# gradient fill | chart plot area background | two color gradient Aspose | horizontal gradient chart .NET | SetTwoColorGradient example | Excel chart styling Aspose.Cells | plot area fill format
// Common Searches: Aspose.Cells set gradient on chart plot area | C# example horizontal gradient chart background | How to use SetTwoColorGradient in Aspose.Cells | Apply complementary colors to Excel chart area .NET | Gradient fill for column chart plot area
// Developer Intent: Add a horizontal two‑color gradient to the plot area of a column chart using Aspose.Cells for .NET.
// Use Cases: Give financial column charts a depth effect with a subtle sky‑to‑navy gradient. | Match corporate branding by applying a custom gradient to chart backgrounds in automated reports. | Visually separate multiple charts in a single workbook by assigning each a distinct gradient style.
// AI Prompts: Write C# code with Aspose.Cells that applies a vertical two‑color gradient from LightGreen to DarkGreen on a line chart's plot area. | Show how to change the gradient variant and direction for a pie chart plot area using Aspose.Cells. | Explain how to read, modify, or replace an existing gradient fill of a chart's plot area after loading an existing workbook.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace GradientPlotAreaDemo
{
    // Demonstrates how to create a workbook, add a column chart, and use the PlotArea.FillFormat.SetTwoColorGradient method to paint the plot area with a LightSkyBlue‑to‑DarkBlue horizontal gradient before saving the file as XLSX.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("Q1");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["A3"].PutValue("Q2");
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["A4"].PutValue("Q3");
            sheet.Cells["B4"].PutValue(90);
            sheet.Cells["A5"].PutValue("Q4");
            sheet.Cells["B5"].PutValue(130);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 15);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B5", true);
            chart.NSeries.CategoryData = "A2:A5";

            // Access the plot area fill format
            FillFormat plotAreaFill = chart.PlotArea.Area.FillFormat;

            // Apply a two‑color gradient (complementary colors) to the plot area
            // Using LightSkyBlue and DarkBlue with a horizontal gradient style
            plotAreaFill.SetTwoColorGradient(
                Color.LightSkyBlue,   // First gradient color
                Color.DarkBlue,       // Second gradient color
                GradientStyleType.Horizontal, // Gradient direction
                1);                   // Variant (1‑4)

            // Save the workbook
            workbook.Save("GradientPlotAreaDemo.xlsx");
        }
    }
}
