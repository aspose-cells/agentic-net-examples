// Title: Apply Built‑In Chart Style and Workbook Theme Colors to a Column Chart with Aspose.Cells for .NET
// Description: Creates a new workbook, adds sample data, inserts a column chart, applies a built‑in chart style (Style 1‑48) via the Chart.Style property, optionally customizes workbook theme colors, and saves the file as an Excel workbook with a consistently themed chart.
// Keywords: Aspose.Cells | C# chart style | built‑in chart theme | Chart.Style property | Workbook theme colors | column chart Aspose.Cells | Excel automation .NET | ThemeColorType | apply chart theme programmatically | standardize chart colors
// Common Searches: how to set a built‑in chart style in Aspose.Cells C# | apply workbook theme colors to Excel charts using Aspose.Cells | change column chart colors with Chart.Style in .NET | list of built‑in chart styles Aspose.Cells | programmatically theme Excel charts Aspose
// Developer Intent: The developer wants to programmatically apply a predefined chart style and optional workbook theme colors to a column chart so that the visual appearance is consistent and matches a corporate palette.
// Use Cases: Generate reports where all charts automatically follow a corporate color scheme without manual formatting. | Switch chart aesthetics across multiple workbooks by changing a single style index or theme color. | Create branded Excel files that enforce uniform fonts, colors, and chart layouts for downstream users.
// AI Prompts: Write C# code using Aspose.Cells to apply a specific built‑in chart style to a line chart and set custom workbook theme colors. | Show how to enumerate all built‑in chart styles (1‑48) in Aspose.Cells and select one based on a user‑defined rule. | Provide an example that updates existing charts after changing workbook theme colors so the new palette is reflected instantly.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a new workbook, adds sample data, inserts a column chart, applies a built‑in chart style (Style 1‑48) via the Chart.Style property, optionally customizes workbook theme colors, and saves the file as an Excel workbook with a consistently themed chart.
class ApplyBuiltInThemeToChart
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
        chart.NSeries.Add("B2:B4", false);          // Values
        chart.NSeries.CategoryData = "A2:A4";       // Categories

        // Apply a built‑in chart style (range 1‑48) to standardize colors and fonts
        chart.Style = 2; // Choose any style number; here we use style 2 as an example

        // Optionally set workbook theme colors to ensure the chart uses consistent palette
        workbook.SetThemeColor(ThemeColorType.Accent1, Color.FromArgb(91, 155, 213)); // Light blue
        workbook.SetThemeColor(ThemeColorType.Accent2, Color.FromArgb(237, 125, 49)); // Orange
        workbook.SetThemeColor(ThemeColorType.Accent3, Color.FromArgb(165, 165, 165)); // Gray

        // Save the workbook with the themed chart
        workbook.Save("ChartWithBuiltInTheme.xlsx");
    }
}
