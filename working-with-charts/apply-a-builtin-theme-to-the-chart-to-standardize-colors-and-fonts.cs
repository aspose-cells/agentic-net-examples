// Title: Apply a Built‑In Chart Theme and Set an Accent Color with Aspose.Cells for C#/.NET
// Description: Creates a workbook, adds sample data, changes the Accent1 theme color, inserts a column chart, applies a built‑in chart style (e.g., style 2) to unify colors and fonts, and saves the file as ChartWithBuiltInTheme.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | .NET | chart theme | built‑in chart style | chart style numbers | set theme color | Accent1 color | column chart | Excel chart formatting | export to XLSX
// Common Searches: Aspose.Cells apply built‑in chart style C# | change workbook theme accent color Aspose.Cells | list chart style numbers Aspose.Cells .NET | standardize Excel chart colors with Aspose.Cells | set Accent1 theme color for chart Aspose.Cells
// Developer Intent: The developer wants to use a predefined chart style and optionally modify a theme accent so the generated Excel chart follows a consistent visual scheme.
// Use Cases: Generate a column chart from worksheet data and enforce a specific style to match corporate branding. | Adjust the workbook’s Accent1 color before applying the chart style so the chart inherits the custom hue. | Produce Excel reports where all charts share the same fonts and palette without manual formatting.
// AI Prompts: Show code that enumerates all built‑in chart style IDs in Aspose.Cells and selects one based on a condition. | Provide an example that sets the Accent2 theme color after applying a chart style, then updates the chart appearance. | Explain how to combine workbook.SetThemeColor with chart.Style to create a fully branded chart in Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook, adds sample data, changes the Accent1 theme color, inserts a column chart, applies a built‑in chart style (e.g., style 2) to unify colors and fonts, and saves the file as ChartWithBuiltInTheme.xlsx using Aspose.Cells for .NET.
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

        // Optionally adjust some theme colors to ensure consistency
        // Here we set Accent1 to a specific color; other accents can be set similarly
        workbook.SetThemeColor(ThemeColorType.Accent1, Color.FromArgb(0, 112, 192)); // a shade of blue

        // Add a column chart
        int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 8);
        Chart chart = sheet.Charts[chartIndex];

        // Define the data range for the chart
        chart.NSeries.Add("B2:B4", false);
        chart.NSeries.CategoryData = "A2:A4";

        // Apply a built‑in chart style (values 1‑48). This standardizes colors and fonts.
        chart.Style = 2; // Example: style number 2

        // Save the workbook with the themed chart
        workbook.Save("ChartWithBuiltInTheme.xlsx");
    }
}
