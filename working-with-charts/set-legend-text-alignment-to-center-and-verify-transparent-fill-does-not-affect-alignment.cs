// Title: Center legend text horizontally and vertically in an Aspose.Cells column chart while applying a transparent background (C#)
// AI Prompts: Generate a column chart in a new workbook, set the legend's TextHorizontalAlignment and TextVerticalAlignment to Center, and apply BackgroundMode.Transparent using Aspose.Cells for .NET. | Add verification code that confirms the legend's alignment properties stay Center after the transparent background is set and output the result to the console.
// Common Searches: Aspose.Cells C# set legend text alignment to center for both horizontal and vertical | make chart legend background transparent without changing alignment Aspose.Cells | verify legend alignment after changing background mode in Aspose.Cells chart | center legend text in column chart using Aspose.Cells for .NET | Aspose.Cells example transparent legend background C#
// Tags: legend text alignment center Aspose.Cells | transparent legend background Aspose.Cells | column chart legend formatting C# | verify legend alignment after background change Aspose.Cells | chart legend properties Aspose.Cells .NET

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsLegendAlignmentDemo
{
    // The sample creates a workbook, adds sample data, builds a column chart, centers the legend's horizontal and vertical text alignment, sets the legend's background to transparent, verifies that the alignment remains centered, prints the verification results, and saves the file as LegendAlignmentTransparentDemo.xlsx.
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
            sheet.Cells["B2"].PutValue(50);
            sheet.Cells["B3"].PutValue(80);
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Access the legend and configure alignment
            Legend legend = chart.Legend;
            legend.TextHorizontalAlignment = TextAlignmentType.Center; // Center horizontally
            legend.TextVerticalAlignment   = TextAlignmentType.Center; // Center vertically

            // Set transparent background to verify it does not affect alignment
            legend.BackgroundMode = BackgroundMode.Transparent;

            // Simple verification that alignment properties remain centered
            bool alignmentIntact = legend.TextHorizontalAlignment == TextAlignmentType.Center &&
                                   legend.TextVerticalAlignment == TextAlignmentType.Center;

            Console.WriteLine("Legend alignment centered: " + alignmentIntact);
            Console.WriteLine("Legend background is transparent: " + (legend.BackgroundMode == BackgroundMode.Transparent));

            // Save the workbook
            workbook.Save("LegendAlignmentTransparentDemo.xlsx");
        }
    }
}
