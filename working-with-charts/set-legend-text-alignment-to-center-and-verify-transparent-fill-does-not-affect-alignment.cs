// Title: Center Chart Legend Text and Preserve Alignment with Transparent Background – Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook, add a column chart, set the legend's TextHorizontalAlignment to Center, apply a Transparent background, verify the alignment remains unchanged, and save the file as LegendAlignmentTransparentDemo.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# chart legend alignment | center legend text Aspose.Cells | transparent legend background Aspose.Cells | BackgroundMode.Transparent chart legend | verify legend alignment .NET | Excel chart legend formatting | Aspose.Cells legend TextHorizontalAlignment | column chart legend styling | Aspose.Cells API legend properties | C# Excel chart example
// Common Searches: how to center legend text in Aspose.Cells chart | Aspose.Cells set legend alignment C# | transparent background for chart legend Aspose.Cells | does BackgroundMode.Transparent affect legend alignment | Aspose.Cells verify legend text alignment after fill | C# code to format chart legend in Aspose.Cells
// Developer Intent: Set a chart legend’s text alignment to Center and ensure that applying a Transparent background does not modify the alignment.
// Use Cases: Design a business dashboard where the legend must be centered for a clean visual layout. | Write automated tests that change the legend’s background to Transparent and assert that TextHorizontalAlignment stays Center. | Generate Excel reports programmatically where legend styling must remain consistent across different fill settings.
// AI Prompts: Create C# code with Aspose.Cells that aligns the chart legend text to the right, applies a solid fill color, and validates the alignment. | Explain the effect of BackgroundMode.Transparent on chart legend rendering in Aspose.Cells and show how to programmatically confirm alignment after changing fill properties. | Provide a unit‑test snippet in C# that sets a legend’s background to Transparent, checks TextHorizontalAlignment, and fails if the alignment is altered.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// Demonstrates how to create a workbook, add a column chart, set the legend's TextHorizontalAlignment to Center, apply a Transparent background, verify the alignment remains unchanged, and save the file as LegendAlignmentTransparentDemo.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
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

        // Add a column chart and set its data range
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
        Chart chart = sheet.Charts[chartIndex];
        chart.SetChartDataRange("A1:B4", true);

        // Access the legend of the chart
        Legend legend = chart.Legend;

        // Set the legend text horizontal alignment to Center
        legend.TextHorizontalAlignment = TextAlignmentType.Center;

        // Apply a transparent background fill to the legend
        legend.BackgroundMode = BackgroundMode.Transparent;

        // Verify that the alignment is still Center after applying transparent fill
        // (simple runtime check; in a real test you would assert this condition)
        if (legend.TextHorizontalAlignment == TextAlignmentType.Center)
        {
            Console.WriteLine("Legend text alignment remains centered despite transparent fill.");
        }
        else
        {
            Console.WriteLine("Legend text alignment was altered.");
        }

        // Save the workbook (lifecycle rule)
        workbook.Save("LegendAlignmentTransparentDemo.xlsx");
    }
}
