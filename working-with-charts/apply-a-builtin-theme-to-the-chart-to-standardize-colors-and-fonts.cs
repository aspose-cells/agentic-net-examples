// Title: Apply Built‑In Chart Theme & Change Workbook Theme Colors – Aspose.Cells C# Example
// Description: C# sample that creates a workbook, fills it with category/value data, adds a column chart, assigns the data range, applies a built‑in chart style (1‑48), optionally updates the workbook’s Accent1 and Accent2 theme colors, and saves the file as ChartWithBuiltInTheme.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# chart style | built‑in chart theme | SetThemeColor | workbook theme colors | column chart example | .NET Excel chart formatting | chart.Style property | Excel theme customization
// Common Searches: Aspose.Cells apply built‑in chart style C# | Change workbook theme colors Aspose.Cells .NET | Set chart theme number 1‑48 Aspose.Cells | How to modify Accent1 Accent2 colors in Excel with Aspose | C# example for chart formatting using Aspose.Cells
// Developer Intent: Apply a predefined chart style and optionally customize the workbook’s theme colors in a .NET Excel file.
// Use Cases: Ensure consistent visual branding across generated reports by setting chart.Style to a specific built‑in theme. | Match corporate color schemes by updating Accent1 and Accent2 via workbook.SetThemeColor after applying the chart style. | Automate creation of column charts from data ranges with standardized appearance for dashboards or exports.
// AI Prompts: Generate C# code that creates a column chart with Aspose.Cells, applies a built‑in chart style (1‑48), changes Accent1 and Accent2 theme colors, and saves the workbook. | Explain how to list all available built‑in chart styles in Aspose.Cells and select one based on a configuration value. | Show a step‑by‑step tutorial for applying a chart theme and customizing workbook theme colors using Aspose.Cells for .NET.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// C# sample that creates a workbook, fills it with category/value data, adds a column chart, assigns the data range, applies a built‑in chart style (1‑48), optionally updates the workbook’s Accent1 and Accent2 theme colors, and saves the file as ChartWithBuiltInTheme.xlsx using Aspose.Cells for .NET.
class ApplyBuiltInChartTheme
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["A4"].PutValue("C");
        sheet.Cells["B4"].PutValue(30);

        // Add a column chart
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = sheet.Charts[chartIndex];

        // Set data range for the chart
        chart.NSeries.Add("B2:B4", false);
        chart.NSeries.CategoryData = "A2:A4";

        // Apply a built‑in style (1‑48). This standardizes colors and fonts.
        chart.Style = 5; // Example style number; choose any between 1 and 48

        // Optionally adjust theme colors globally (e.g., change Accent1 and Accent2)
        workbook.SetThemeColor(ThemeColorType.Accent1, Color.FromArgb(0, 112, 192)); // Dark blue
        workbook.SetThemeColor(ThemeColorType.Accent2, Color.FromArgb(255, 192, 0)); // Orange

        // Save the workbook
        workbook.Save("ChartWithBuiltInTheme.xlsx");
    }
}
