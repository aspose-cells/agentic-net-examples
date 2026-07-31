// Title: Aspose.Cells for .NET – Set Chart Legend Font to Dark Gray and Make Background Transparent (C#)
// Description: Creates a workbook, adds a column chart, displays the legend, and configures each LegendEntry so the text color is DarkGray while the legend background is transparent yet still readable. The workbook is saved as an .xlsx file.
// Keywords: Aspose.Cells legend font color | chart legend dark gray C# | transparent legend background Aspose.Cells | set legend entry style .NET | Excel chart legend readability | Aspose.Cells chart formatting | C# Aspose.Cells legend customization
// Common Searches: change legend text color Aspose.Cells | transparent chart legend background .NET | set dark gray legend font Aspose.Cells C# | make legend background invisible Excel chart | Aspose.Cells legend entry formatting example
// Developer Intent: Apply a dark‑gray font to all legend entries and render the legend background transparent without compromising text visibility.
// Use Cases: Design corporate reports where the legend must blend with patterned worksheet backgrounds while keeping text legible. | Generate dashboards that overlay charts on colored cells, requiring a transparent legend to avoid obscuring underlying data. | Standardize legend appearance across multiple charts in a workbook for consistent branding and readability.
// AI Prompts: Write C# code using Aspose.Cells to set legend entry font to a specific RGB value and make the legend background transparent while preserving text fill. | Explain how BackgroundMode.Transparent and IsTextNoFill affect legend rendering in Aspose.Cells charts. | Provide a method to apply the same legend font color and transparency settings to every chart in an Aspose.Cells workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

// Creates a workbook, adds a column chart, displays the legend, and configures each LegendEntry so the text color is DarkGray while the legend background is transparent yet still readable. The workbook is saved as an .xlsx file.
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
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(50);
        sheet.Cells["B3"].PutValue(80);

        // Add a column chart
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
        Chart chart = sheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B3", true);          // Values
        chart.NSeries.CategoryData = "A2:A3";      // Categories
        chart.ShowLegend = true;                   // Ensure legend is displayed

        // Configure each legend entry
        foreach (LegendEntry entry in chart.Legend.LegendEntries)
        {
            // Set the legend text color to dark gray
            entry.Font.Color = Color.DarkGray;

            // Make the legend background transparent but keep text fill
            entry.BackgroundMode = BackgroundMode.Transparent;
            entry.IsTextNoFill = false; // Ensure text has a fill for readability
        }

        // Save the workbook
        workbook.Save("LegendFontDarkGray.xlsx");
    }
}
