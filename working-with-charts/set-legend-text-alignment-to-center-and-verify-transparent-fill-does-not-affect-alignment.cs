// Title: Center Legend Text in Aspose.Cells Chart (C#) with Transparent Background
// Description: Creates a workbook, adds a column chart, centers the legend text horizontally and vertically, sets the legend background to transparent to confirm alignment is unchanged, and saves the file as an Excel workbook.
// Keywords: Aspose.Cells | C# chart legend alignment | center legend text | transparent legend background | legend text alignment .NET | Excel chart legend | Aspose.Cells legend properties
// Common Searches: Aspose.Cells center legend text | how to align chart legend horizontally and vertically in C# | transparent legend background Aspose.Cells | does legend background affect alignment in Aspose.Cells charts | C# example for chart legend alignment Aspose.Cells
// Developer Intent: Align chart legend text to the center on both axes while using a transparent background to ensure alignment remains consistent.
// Use Cases: Generate reports with column charts that have a centrally aligned legend for a clean visual layout. | Apply a transparent legend background when the chart overlays colored cells, preserving text positioning. | Create reusable chart templates where legend alignment is predefined, simplifying automated Excel generation.
// AI Prompts: Write C# code using Aspose.Cells to add a chart and set the legend text alignment to center with a transparent background. | Explain how a transparent legend background impacts (or does not impact) text alignment in Aspose.Cells charts and how to test it. | Provide step‑by‑step instructions to center legend text horizontally and vertically and verify alignment after changing the legend's background mode.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsLegendAlignmentDemo
{
    // Creates a workbook, adds a column chart, centers the legend text horizontally and vertically, sets the legend background to transparent to confirm alignment is unchanged, and saves the file as an Excel workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Q1");
            sheet.Cells["A3"].PutValue("Q2");
            sheet.Cells["A4"].PutValue("Q3");

            sheet.Cells["B1"].PutValue("Series1");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["B4"].PutValue(180);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Access the legend and set its properties
            Legend legend = chart.Legend;

            // Center the legend text horizontally and vertically
            legend.TextHorizontalAlignment = TextAlignmentType.Center;
            legend.TextVerticalAlignment = TextAlignmentType.Center;

            // Make the legend background transparent to verify it does not affect alignment
            legend.BackgroundMode = BackgroundMode.Transparent;

            // Save the workbook
            workbook.Save("LegendAlignmentTransparentDemo.xlsx");
        }
    }
}
