// Title: Aspose.Cells .NET – Change Theme Color of the Second Series in a Pie Chart
// Description: Creates a workbook, adds two data series to a pie chart, then uses a Style object to set a custom ForegroundColor and applies it to the Area of the second series (index 1) before saving the file.
// Keywords: Aspose.Cells change series color | pie chart second series color .NET | chart series style Aspose.Cells | Area.ForegroundColor Aspose | custom theme color chart series
// Common Searches: how to set a different color for the second series in an Aspose.Cells pie chart | Aspose.Cells change theme color of a specific chart series | apply custom style to chart series area C# | Aspose.Cells modify series fill color without affecting whole chart
// Developer Intent: Apply a custom theme color to only the second series of a pie chart using Aspose.Cells for .NET.
// Use Cases: Highlight a secondary data series with a brand‑specific color in automated report charts. | Enforce distinct colors for individual series when generating multi‑series pie charts programmatically. | Create visual contrast between series without altering the chart’s global theme.
// AI Prompts: Generate C# code with Aspose.Cells to set a custom foreground color for the third series of a bar chart. | Show how to apply a gradient Style object to a chart series area in Aspose.Cells. | Explain how to access and modify the Area properties of a chart series to change its fill and border colors.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// Creates a workbook, adds two data series to a pie chart, then uses a Style object to set a custom ForegroundColor and applies it to the Area of the second series (index 1) before saving the file.
class ChangeSecondSeriesThemeColor
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate data for two series in a pie chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("Apple");
        sheet.Cells["A3"].PutValue("Orange");
        sheet.Cells["A4"].PutValue("Banana");

        // First series values
        sheet.Cells["B1"].PutValue("Series1");
        sheet.Cells["B2"].PutValue(50);
        sheet.Cells["B3"].PutValue(30);
        sheet.Cells["B4"].PutValue(20);

        // Second series values (to demonstrate separate styling)
        sheet.Cells["C1"].PutValue("Series2");
        sheet.Cells["C2"].PutValue(40);
        sheet.Cells["C3"].PutValue(35);
        sheet.Cells["C4"].PutValue(25);

        // Add a pie chart
        int chartIdx = sheet.Charts.Add(ChartType.Pie, 5, 0, 20, 12);
        Chart chart = sheet.Charts[chartIdx];

        // Add both series to the chart
        // First series (B2:B4)
        chart.NSeries.Add("B2:B4", true);
        // Second series (C2:C4)
        chart.NSeries.Add("C2:C4", true);

        // Set category data (A2:A4)
        chart.NSeries.CategoryData = "A2:A4";

        // -------------------------------------------------
        // Change the theme color of the second series
        // -------------------------------------------------
        // Create a new Style object
        Style newStyle = workbook.CreateStyle();

        // Assign a foreground color to the style (this will act as the theme color)
        newStyle.ForegroundColor = Color.FromArgb(255, 102, 0); // Example: a vivid orange

        // Apply the style's foreground color to the second series area
        // (Series index 1 corresponds to the second series)
        chart.NSeries[1].Area.ForegroundColor = newStyle.ForegroundColor;

        // Save the workbook
        workbook.Save("PieChart_SecondSeriesThemeColor.xlsx");
    }
}
