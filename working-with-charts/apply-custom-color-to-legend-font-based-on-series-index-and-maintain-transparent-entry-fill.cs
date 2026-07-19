// Title: Set per‑series legend font colors with transparent background in Aspose.Cells C# chart
// Description: Creates a workbook, adds a column chart with two series, and assigns each legend entry a font color from a palette while keeping the legend background transparent, then saves the file.
// Keywords: Aspose.Cells legend font color | C# chart legend custom colors | transparent legend background Aspose | per series legend formatting | Aspose.Cells chart styling
// Common Searches: Aspose.Cells change legend text color per series | how to make legend background transparent in Aspose.Cells | C# set custom colors for chart legend entries | Aspose.Cells legend formatting examples
// Developer Intent: Apply a distinct font color to each legend entry based on its series index while preserving a transparent legend background.
// Use Cases: Generate a multi‑series column chart where each legend label matches a brand color palette. | Programmatically style legends for dashboards that overlay chart colors on varied worksheet backgrounds. | Loop through an arbitrary number of series and reuse a limited set of colors without affecting legend fill.
// AI Prompts: Write C# code using Aspose.Cells to color legend text per series and keep the legend background transparent. | Show how to extend the legend‑color loop for any number of series with a repeating color array. | Explain the impact of LegendEntry.BackgroundMode = Transparent on chart rendering in Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook, adds a column chart with two series, and assigns each legend entry a font color from a palette while keeping the legend background transparent, then saves the file.
class CustomLegendColors
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
        sheet.Cells["B2"].PutValue(120);
        sheet.Cells["B3"].PutValue(150);
        sheet.Cells["B4"].PutValue(180);

        sheet.Cells["C1"].PutValue("Series 2");
        sheet.Cells["C2"].PutValue(80);
        sheet.Cells["C3"].PutValue(130);
        sheet.Cells["C4"].PutValue(170);

        // Add a column chart
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
        Chart chart = sheet.Charts[chartIndex];

        // Add two series to the chart
        chart.NSeries.Add("B2:B4", true); // Series 1
        chart.NSeries.Add("C2:C4", true); // Series 2
        chart.NSeries.CategoryData = "A2:A4";

        // Define custom colors for legend entries (one per series)
        Color[] legendColors = new Color[]
        {
            Color.FromArgb(79, 129, 189),   // Color for series 0
            Color.FromArgb(192, 80, 77)    // Color for series 1
        };

        // Apply custom font color and keep the legend entry background transparent
        for (int i = 0; i < chart.NSeries.Count; i++)
        {
            LegendEntry entry = chart.NSeries[i].LegendEntry;

            // Set the font color based on the series index
            entry.Font.Color = legendColors[i % legendColors.Length];

            // Ensure the legend entry background remains transparent
            entry.BackgroundMode = BackgroundMode.Transparent;
        }

        // Save the workbook
        workbook.Save("CustomLegendColors.xlsx");
    }
}
