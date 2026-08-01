// Title: Aspose.Cells for .NET: Change Theme Color of Second Series in a Pie Chart (C#)
// Description: Creates a workbook, adds two data series to a pie chart, defines a Style with a ThemeColor (Accent2), applies that ThemeColor to the border of the second series, and saves the workbook.
// Keywords: Aspose.Cells | C# chart theme color | pie chart series style | Series.Border.ThemeColor | Style object | ThemeColor Accent2 | Aspose.Cells chart customization | set series border color | Aspose.Cells .NET example
// Common Searches: Aspose.Cells set theme color for chart series | Change pie chart series border color C# | Apply Style ThemeColor to chart series Aspose.Cells | How to use Style object to color chart series Aspose | Aspose.Cells change second series color in pie chart
// Developer Intent: Apply a ThemeColor via a Style object to the border of the second series in a pie chart using Aspose.Cells for .NET.
// Use Cases: Match chart series colors to corporate theme (e.g., Accent2) for branding consistency. | Highlight a specific series in a multi‑series pie chart by altering its border color. | Reuse a predefined Style with a ThemeColor across multiple charts for uniform appearance. | Generate reports where each series is distinguished by a distinct theme color.
// AI Prompts: Write C# code to set the ThemeColor of the third series in a bar chart using Aspose.Cells and a Style object. | Show how to apply a custom ThemeColor to data labels of a chart series in Aspose.Cells. | Create an example that assigns different line styles and ThemeColors to each series in a stacked column chart with Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using System.Drawing;

namespace AsposeCellsThemeColorDemo
{
    // Creates a workbook, adds two data series to a pie chart, defines a Style with a ThemeColor (Accent2), applies that ThemeColor to the border of the second series, and saves the workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle: create)
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

            // Second series values
            sheet.Cells["C1"].PutValue("Series2");
            sheet.Cells["C2"].PutValue(40);
            sheet.Cells["C3"].PutValue(35);
            sheet.Cells["C4"].PutValue(25);

            // Add a pie chart
            int chartIdx = sheet.Charts.Add(ChartType.Pie, 5, 0, 20, 12);
            Chart chart = sheet.Charts[chartIdx];

            // Add two series to the chart
            chart.NSeries.Add("B2:B4", true); // first series
            chart.NSeries.Add("C2:C4", true); // second series
            chart.NSeries.CategoryData = "A2:A4";

            // ------------------------------------------------------------
            // Change the theme color of the second series (index 1)
            // ------------------------------------------------------------

            // Create a new Style object (as requested)
            Style newStyle = workbook.CreateStyle();

            // Set a theme color for the font in the style (example: Accent2)
            // This demonstrates creating a style with a theme color.
            newStyle.Font.ThemeColor = new ThemeColor(ThemeColorType.Accent2, 0.0);

            // Apply the theme color to the second series' border.
            // The Series.Border is a Line object which supports ThemeColor.
            Series secondSeries = chart.NSeries[1];
            secondSeries.Border.ThemeColor = newStyle.Font.ThemeColor;

            // Optionally, set the border style and weight for better visibility
            secondSeries.Border.Style = LineType.Solid;
            secondSeries.Border.Weight = WeightType.MediumLine;

            // Save the workbook (lifecycle: save)
            workbook.Save("PieChart_SecondSeries_ThemeColor.xlsx");
        }
    }
}
