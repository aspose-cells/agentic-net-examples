using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

class ChangeChartToOfficeTheme
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate some sample data for the chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["A4"].PutValue("C");
        sheet.Cells["B4"].PutValue(30);

        // Define the 12 colors of the built‑in "Office" theme
        Color[] officeThemeColors = new Color[]
        {
            Color.White,                     // Background1
            Color.Black,                     // Text1
            Color.White,                     // Background2
            Color.Black,                     // Text2
            Color.FromArgb(0, 112, 192),     // Accent1 (blue)
            Color.FromArgb(255, 192, 0),     // Accent2 (orange)
            Color.FromArgb(112, 48, 160),    // Accent3 (purple)
            Color.FromArgb(0, 176, 80),      // Accent4 (green)
            Color.FromArgb(255, 0, 0),       // Accent5 (red)
            Color.FromArgb(255, 0, 255),     // Accent6 (magenta)
            Color.FromArgb(0, 0, 255),       // Hyperlink (blue)
            Color.FromArgb(128, 0, 128)      // Followed Hyperlink (purple)
        };

        // Apply the "Office" theme to the workbook
        workbook.CustomTheme("Office", officeThemeColors);

        // Add a column chart that will inherit the theme colors
        int chartIdx = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 10);
        Chart chart = sheet.Charts[chartIdx];
        chart.NSeries.Add("B2:B4", false);
        chart.NSeries.CategoryData = "A2:A4";

        // Optionally set a built‑in chart style (1‑48) – here we keep default
        // chart.Style = 1;

        // Save the workbook; the chart now uses the Office theme colors
        workbook.Save("ChartWithOfficeTheme.xlsx");
    }
}