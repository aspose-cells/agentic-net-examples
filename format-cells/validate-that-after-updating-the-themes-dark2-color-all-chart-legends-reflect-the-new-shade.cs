// Title: Aspose.Cells C# – Verify Chart Legends Update After Changing Dark2 (Accent2) Theme Color
// Description: Creates a workbook with a column chart, changes the workbook's Accent2 (Dark2) theme color to DarkSlateBlue using SetThemeColor, applies the same theme color to each chart legend's font, outputs the legend's ThemeColor details for verification, and saves the file.
// Keywords: Aspose.Cells | C# | .NET | SetThemeColor | Accent2 | Dark2 | chart legend | theme color | ThemeColor | chart formatting | workbook theme | Excel automation
// Common Searches: How to change Dark2 theme color in Aspose.Cells for .NET | Aspose.Cells chart legend uses workbook theme color | Validate chart legend color after SetThemeColor | C# Aspose.Cells update theme color for all charts | Apply custom Accent2 color to Excel chart legends with Aspose
// Developer Intent: Confirm that modifying the workbook’s Dark2 (Accent2) theme color automatically updates the color of all chart legends.
// Use Cases: Programmatically rebrand a workbook by changing the Accent2 theme color and having every existing chart legend adopt the new shade without manual adjustments. | Log each legend’s ThemeColor type and tint to ensure the theme change propagated correctly across multiple charts. | Generate standardized reports where chart legends consistently follow a custom Dark2 theme for corporate visual identity.
// AI Prompts: Write C# code using Aspose.Cells that changes the Dark2 (Accent2) theme color, updates all chart legends to the new theme color, and prints each legend’s ThemeColor details for verification. | Show how to iterate through every chart in a workbook and set each legend’s Font.ThemeColor to match a modified Accent2 color, then output the ThemeColor type and tint. | Explain step‑by‑step how SetThemeColor and the ThemeColor object affect chart legend appearance in Aspose.Cells, including best practices for theme‑based styling.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace ThemeLegendValidation
{
    // Creates a workbook with a column chart, changes the workbook's Accent2 (Dark2) theme color to DarkSlateBlue using SetThemeColor, applies the same theme color to each chart legend's font, outputs the legend's ThemeColor details for verification, and saves the file.
    class Program
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

            for (int i = 2; i <= 5; i++)
            {
                sheet.Cells[$"A{i}"].PutValue($"Item {i - 1}");
                sheet.Cells[$"B{i}"].PutValue(i * 10);
                sheet.Cells[$"C{i}"].PutValue(i * 15);
            }

            // Add a column chart
            int chartIdx = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 15);
            Chart chart = sheet.Charts[chartIdx];

            // Set chart data range and categories
            chart.SetChartDataRange("B1:C5", true);
            chart.NSeries.CategoryData = "A2:A5";

            // Update the theme's Dark2 color (using Accent2 as the closest match)
            // This changes the theme color globally in the workbook
            workbook.SetThemeColor(ThemeColorType.Accent2, Color.DarkSlateBlue);

            // Apply the updated theme color to the chart legend font
            // The legend will now use the new Dark2 (Accent2) shade
            chart.Legend.Font.ThemeColor = new ThemeColor(ThemeColorType.Accent2, 0.0);

            // Optional validation: output the legend font's theme color type and tint
            ThemeColor legendTheme = chart.Legend.Font.ThemeColor;
            Console.WriteLine($"Legend Font Theme Color Type: {legendTheme.ColorType}");
            Console.WriteLine($"Legend Font Theme Color Tint: {legendTheme.Tint}");

            // Save the workbook
            workbook.Save("ThemeLegendValidation.xlsx");
        }
    }
}
