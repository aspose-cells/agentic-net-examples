// Title: Aspose.Cells C# – Set Chart Legend Font to Dark Gray with Transparent Background
// Description: Create a workbook, add a column chart, display the legend, and programmatically set each legend entry's font color to DarkGray while keeping the legend background transparent and the text readable using Aspose.Cells for .NET.
// Keywords: Aspose.Cells legend font color | chart legend dark gray C# | transparent legend background Aspose | LegendEntry formatting Aspose.Cells | C# Aspose.Cells chart customization | set legend text fill Aspose | Aspose.Cells chart styling
// Common Searches: how to change legend font color in Aspose.Cells | transparent chart legend Aspose.Cells C# | set legend entry font color dark gray Aspose | Aspose.Cells legend background transparent | customize chart legend text fill Aspose
// Developer Intent: The developer needs to modify a chart's legend so the text appears in dark gray while the legend background stays transparent, ensuring the text remains legible.
// Use Cases: Standardize legend appearance across corporate reports generated with Aspose.Cells. | Overlay charts on colored slides or dashboards where a transparent legend improves visual integration. | Apply consistent legend styling to multiple charts in a workbook via automated code.
// AI Prompts: Generate C# code with Aspose.Cells that sets each legend entry's font color to DarkGray and makes the legend background transparent without losing readability. | Show how to loop through Chart.Legend.LegendEntries to apply Font.Color, BackgroundMode, and IsTextNoFill settings in Aspose.Cells. | Explain the role of LegendEntry.IsTextNoFill and BackgroundMode.Transparent for readable transparent legends in Aspose.Cells charts.

using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Create a workbook, add a column chart, display the legend, and programmatically set each legend entry's font color to DarkGray while keeping the legend background transparent and the text readable using Aspose.Cells for .NET.
class SetLegendFontColor
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
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
        Chart chart = sheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Ensure the legend is displayed
        chart.ShowLegend = true;

        // Iterate through each legend entry and set the desired properties
        foreach (LegendEntry entry in chart.Legend.LegendEntries)
        {
            // Set font color to dark gray
            entry.Font.Color = Color.DarkGray;

            // Keep background transparent but ensure text fill is applied
            entry.BackgroundMode = BackgroundMode.Transparent;
            entry.IsTextNoFill = false; // text will have fill, preserving readability
        }

        // Save the workbook
        workbook.Save("LegendFontDarkGray.xlsx");
    }
}
