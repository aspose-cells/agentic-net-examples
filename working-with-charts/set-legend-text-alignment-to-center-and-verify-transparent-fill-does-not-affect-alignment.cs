// Title: Center Chart Legend Text and Apply Transparent Background with Aspose.Cells (C#)
// Description: The sample builds a workbook, inserts a column chart, and shows how to set the legend's horizontal and vertical text alignment to Center while using a Transparent background. It also guarantees the legend stays visible and saves the result as LegendAlignmentTransparentFill.xlsx.
// Keywords: Aspose.Cells | C# chart legend alignment | center legend text | transparent legend background | TextHorizontalAlignment | TextVerticalAlignment | BackgroundMode.Transparent | column chart generation | Excel file creation .NET | chart formatting Aspose.Cells
// Common Searches: Aspose.Cells set legend alignment C# | transparent legend background Aspose.Cells chart | center legend horizontally and vertically .NET | verify legend positioning after background change | chart legend formatting example Aspose.Cells
// Developer Intent: Align the legend text to the middle of the box and ensure a transparent fill does not alter that positioning.
// Use Cases: Produce reports where the legend appears centered for a balanced visual layout. | Overlay charts on colored sheets without a solid legend box, keeping text centered. | Automate Excel generation where legend alignment must remain stable regardless of fill settings.
// AI Prompts: Generate C# code that centers both the horizontal and vertical text of a chart legend in Aspose.Cells and saves the workbook. | Show how to set the legend's background to Transparent in Aspose.Cells while preserving its alignment. | Explain steps to test that legend alignment stays unchanged after applying a transparent fill in an Aspose.Cells chart.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsLegendAlignmentDemo
{
    // The sample builds a workbook, inserts a column chart, and shows how to set the legend's horizontal and vertical text alignment to Center while using a Transparent background. It also guarantees the legend stays visible and saves the result as LegendAlignmentTransparentFill.xlsx.
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

            sheet.Cells["B1"].PutValue("Series 1");
            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["B3"].PutValue(50);
            sheet.Cells["B4"].PutValue(70);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Access the legend and set its properties
            Legend legend = chart.Legend;

            // Center the legend text horizontally and vertically
            legend.TextHorizontalAlignment = TextAlignmentType.Center;
            legend.TextVerticalAlignment = TextAlignmentType.Center;

            // Set the legend background to transparent to verify it does not affect alignment
            legend.BackgroundMode = BackgroundMode.Transparent;

            // Optionally, ensure the legend is visible
            chart.ShowLegend = true;

            // Save the workbook
            workbook.Save("LegendAlignmentTransparentFill.xlsx");
        }
    }
}
