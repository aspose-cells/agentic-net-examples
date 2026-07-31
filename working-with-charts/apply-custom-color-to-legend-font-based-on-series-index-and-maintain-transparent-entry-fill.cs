// Title: Set Legend Font Color per Series with Transparent Background in Aspose.Cells .NET
// Description: Shows how to create a workbook, add a column chart with multiple series, and use Aspose.Cells for .NET to assign a distinct font color to each legend entry based on its series index while keeping the legend background transparent, then save the file as XLSX.
// Keywords: Aspose.Cells | C# chart legend formatting | custom legend font color | transparent legend background | series index color | column chart Aspose.Cells | .NET Excel automation | legend entry styling
// Common Searches: Aspose.Cells change legend font color per series | transparent legend entry background Aspose.Cells | set legend entry font color C# | custom legend colors Aspose.Cells chart | how to make legend background transparent in Excel using Aspose
// Developer Intent: Apply a unique font color to each legend entry according to its series order and keep the legend entry background transparent.
// Use Cases: Design a multi‑series column chart where each legend label uses a different color for quick visual identification. | Produce reports that require the legend to blend seamlessly with the worksheet background, necessitating a transparent fill. | Cycle through a predefined palette when the number of series exceeds available colors, ensuring consistent styling. | Create dashboards where legend text color matches series colors while the background remains unobtrusive.
// AI Prompts: Generate C# code that applies a gradient of font colors to legend entries instead of a fixed color array using Aspose.Cells. | Explain how to modify legend entry font size, style, and weight together with custom colors in the provided example. | Show how to set a semi‑transparent background (e.g., 50% opacity) for legend entries and retrieve the current background mode.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Shows how to create a workbook, add a column chart with multiple series, and use Aspose.Cells for .NET to assign a distinct font color to each legend entry based on its series index while keeping the legend background transparent, then save the file as XLSX.
class LegendCustomColorExample
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
        sheet.Cells["A5"].PutValue("Q4");

        sheet.Cells["B1"].PutValue("Series 1");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["B4"].PutValue(30);
        sheet.Cells["B5"].PutValue(40);

        sheet.Cells["C1"].PutValue("Series 2");
        sheet.Cells["C2"].PutValue(15);
        sheet.Cells["C3"].PutValue(25);
        sheet.Cells["C4"].PutValue(35);
        sheet.Cells["C5"].PutValue(45);

        // Add a column chart
        int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 20, 15);
        Chart chart = sheet.Charts[chartIndex];

        // Add two series to the chart
        chart.NSeries.Add("B2:B5", true); // Series 1
        chart.NSeries.Add("C2:C5", true); // Series 2
        chart.NSeries.CategoryData = "A2:A5";

        // Define custom colors for legend entries (one per series)
        Color[] legendColors = new Color[] { Color.Red, Color.Green, Color.Blue, Color.Orange };

        // Apply custom font color and transparent background to each legend entry
        for (int i = 0; i < chart.NSeries.Count; i++)
        {
            LegendEntry entry = chart.NSeries[i].LegendEntry;

            // Set transparent background for the legend entry
            entry.BackgroundMode = BackgroundMode.Transparent;

            // Choose a color based on the series index (cycle if more series than colors)
            Color fontColor = legendColors[i % legendColors.Length];
            entry.Font.Color = fontColor;
        }

        // Save the workbook
        workbook.Save("LegendCustomColorExample.xlsx");
    }
}
