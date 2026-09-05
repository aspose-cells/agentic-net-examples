// Title: How to apply a built‑in chart style and custom workbook theme colors to a chart with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code using Aspose.Cells that creates a column chart, defines custom accent colors in the workbook theme, and assigns a built‑in chart style to the chart. | Show how to set the chart title font to use a theme accent color after configuring workbook theme colors in Aspose.Cells for .NET. | Adapt the sample to generate a line chart, change the built‑in style index, and keep the custom theme colors applied to all chart elements.
// Common Searches: aspnet apply built‑in chart style Aspose.Cells example | custom workbook theme colors for charts in Aspose.Cells C# | which chart style numbers are available in Aspose.Cells .NET | change chart title font to theme accent color using Aspose.Cells | how to use theme colors with built‑in chart styles in Aspose.Cells
// Tags: use built‑in chart style Aspose.Cells | define workbook theme accent colors C# | set chart title font to theme color Aspose.Cells | create column chart with custom theme Aspose.Cells | available built‑in chart style numbers Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// The example creates a new workbook, adds sample data, sets custom accent colors in the workbook theme, inserts a column chart bound to the data, applies a built‑in chart style that leverages the theme colors, configures the chart title to use a theme‑based font color, and saves the file as ChartWithBuiltInTheme.xlsx.
class ApplyBuiltInChartTheme
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
            sheet.Cells[$"A{i}"].PutValue("Cat " + (i - 1));
            sheet.Cells[$"B{i}"].PutValue(i * 10);
            sheet.Cells[$"C{i}"].PutValue(i * 15);
        }

        // Optionally adjust some theme colors – these will be used by the chart style
        workbook.SetThemeColor(ThemeColorType.Accent1, Color.FromArgb(0, 112, 192)); // blue accent
        workbook.SetThemeColor(ThemeColorType.Accent2, Color.FromArgb(255, 192, 0)); // orange accent

        // Add a column chart to the worksheet
        int chartIdx = sheet.Charts.Add(ChartType.Column, 8, 0, 22, 10);
        Chart chart = sheet.Charts[chartIdx];

        // Bind the data range to the chart
        chart.NSeries.Add("B2:C6", true);          // values
        chart.NSeries.CategoryData = "A2:A6";      // categories

        // Apply a built‑in chart style (valid values are 1‑48). The style uses the workbook's theme colors.
        chart.Style = 2; // choose any style number you prefer

        // Set chart title and apply a theme color to its font
        chart.Title.Text = "Sales by Category";
        chart.Title.Font.ThemeColor = new ThemeColor(ThemeColorType.Accent1, 0.0);
        chart.Title.Font.Size = 14;
        chart.Title.Font.IsBold = true;

        // Save the workbook with the themed chart
        workbook.Save("ChartWithBuiltInTheme.xlsx");
    }
}
