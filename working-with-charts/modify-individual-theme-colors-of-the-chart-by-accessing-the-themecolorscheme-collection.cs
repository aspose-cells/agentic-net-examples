// Title: Set Individual Chart Theme Colors with Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, add a column chart, and customize its palette by assigning specific ThemeColorType values (Accent1‑Accent6, Text1, Text2) using Workbook.SetThemeColor, which instantly reflects in the chart before saving the file.
// Keywords: Aspose.Cells chart theme colors | Workbook.SetThemeColor .NET | custom Excel chart palette | change Accent1 Accent2 Aspose | modify chart text colors programmatically | C# Aspose.Cells theme color scheme
// Common Searches: how to change chart theme colors Aspose.Cells | set custom accent colors for Excel chart .NET | Workbook.SetThemeColor example code | modify text colors in Aspose chart theme | Aspose.Cells change theme color scheme
// Developer Intent: Apply specific theme colors to a chart by updating the workbook’s ThemeColorType settings.
// Use Cases: Brand a report with a corporate pastel palette by redefining Accent1‑Accent6 colors. | Enhance axis‑label readability by assigning darker Text1/Text2 theme colors after chart creation. | Ensure consistent visual styling across multiple charts in a workbook by setting a single theme color scheme.
// AI Prompts: Write C# code that changes only Accent3 and Accent5 theme colors of an existing Aspose.Cells chart. | Explain how Workbook.SetThemeColor updates chart rendering and which ThemeColorType values affect different chart elements. | Provide a method to revert customized theme colors back to the default Aspose.Cells theme.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Demonstrates how to create a workbook, add a column chart, and customize its palette by assigning specific ThemeColorType values (Accent1‑Accent6, Text1, Text2) using Workbook.SetThemeColor, which instantly reflects in the chart before saving the file.
class ModifyChartThemeColors
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Series1");
        sheet.Cells["C1"].PutValue("Series2");

        for (int i = 2; i <= 6; i++)
        {
            sheet.Cells[$"A{i}"].PutValue($"Cat {i - 1}");
            sheet.Cells[$"B{i}"].PutValue(i * 10);
            sheet.Cells[$"C{i}"].PutValue(i * 15);
        }

        // Add a column chart
        int chartIdx = sheet.Charts.Add(ChartType.Column, 8, 0, 25, 15);
        Chart chart = sheet.Charts[chartIdx];

        // Set the data range for the chart
        chart.NSeries.Add("B2:C6", true);
        chart.NSeries.CategoryData = "A2:A6";

        // -----------------------------------------------------------------
        // Modify individual theme colors that the chart will use.
        // The theme colors are stored in the workbook's theme color scheme.
        // Changing them via SetThemeColor updates the chart automatically.
        // -----------------------------------------------------------------
        workbook.SetThemeColor(ThemeColorType.Accent1, Color.FromArgb(255, 102, 102)); // Light red
        workbook.SetThemeColor(ThemeColorType.Accent2, Color.FromArgb(102, 255, 102)); // Light green
        workbook.SetThemeColor(ThemeColorType.Accent3, Color.FromArgb(102, 102, 255)); // Light blue
        workbook.SetThemeColor(ThemeColorType.Accent4, Color.FromArgb(255, 255, 102)); // Light yellow
        workbook.SetThemeColor(ThemeColorType.Accent5, Color.FromArgb(255, 102, 255)); // Light magenta
        workbook.SetThemeColor(ThemeColorType.Accent6, Color.FromArgb(102, 255, 255)); // Light cyan

        // Optionally, change text colors to see the effect on axis labels, etc.
        workbook.SetThemeColor(ThemeColorType.Text1, Color.DarkSlateGray);
        workbook.SetThemeColor(ThemeColorType.Text2, Color.DimGray);

        // Save the workbook with the modified theme colors applied to the chart
        workbook.Save("ChartWithCustomThemeColors.xlsx");
    }
}
