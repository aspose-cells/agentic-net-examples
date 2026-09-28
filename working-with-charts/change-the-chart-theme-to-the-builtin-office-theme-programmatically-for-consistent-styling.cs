// Title: Programmatically apply the built‑in Office theme to a column chart using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that applies the Office palette via Workbook.SetThemeColor and builds a column chart that inherits those colors using Aspose.Cells. | Show how to set each ThemeColorType to the standard Office palette before creating a chart, then assign a built‑in chart style (e.g., style 2) in Aspose.Cells for .NET. | Provide a step‑by‑step example that creates sample data, configures the Office color palette, adds a column chart, applies a chart style, and saves the workbook as an .xlsx file with Aspose.Cells C# API.
// Common Searches: Aspose.Cells C# set Office theme colors for workbook and chart | How to apply built‑in Office theme to a chart programmatically with Aspose.Cells .NET | Change chart style to match Office theme using Aspose.Cells C# example | Configure ThemeColorType values for default Office palette in Aspose.Cells | Create column chart that inherits workbook theme in Aspose.Cells C#
// Tags: configure workbook theme Aspose.Cells C# | chart inherits workbook theme Aspose.Cells | apply built‑in chart style 2 Aspose.Cells | create column chart with theme colors Aspose.Cells | save workbook as xlsx Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, populates sample data, sets the built‑in Office palette using Workbook.SetThemeColor for each ThemeColorType, adds a column chart that automatically uses those theme colors, applies chart style 2, and saves the file as ChartWithOfficeTheme.xlsx.
class ChangeChartThemeToOffice
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

        // Apply the built‑in "Office" theme colors to the workbook.
        // These colors correspond to the default Office theme.
        workbook.SetThemeColor(ThemeColorType.Background1, Color.White);               // Background1
        workbook.SetThemeColor(ThemeColorType.Text1, Color.Black);                   // Text1
        workbook.SetThemeColor(ThemeColorType.Background2, Color.FromArgb(242, 242, 242)); // Background2 (light gray)
        workbook.SetThemeColor(ThemeColorType.Text2, Color.FromArgb(89, 89, 89));    // Text2 (dark gray)
        workbook.SetThemeColor(ThemeColorType.Accent1, Color.FromArgb(0, 112, 192)); // Accent1 (blue)
        workbook.SetThemeColor(ThemeColorType.Accent2, Color.FromArgb(255, 192, 0)); // Accent2 (orange)
        workbook.SetThemeColor(ThemeColorType.Accent3, Color.FromArgb(112, 173, 71)); // Accent3 (green)
        workbook.SetThemeColor(ThemeColorType.Accent4, Color.FromArgb(255, 0, 0));   // Accent4 (red)
        workbook.SetThemeColor(ThemeColorType.Accent5, Color.FromArgb(0, 176, 80)); // Accent5 (lime)
        workbook.SetThemeColor(ThemeColorType.Accent6, Color.FromArgb(112, 48, 160)); // Accent6 (purple)
        workbook.SetThemeColor(ThemeColorType.Hyperlink, Color.FromArgb(0, 0, 255)); // Hyperlink (blue)
        workbook.SetThemeColor(ThemeColorType.FollowedHyperlink, Color.FromArgb(128, 0, 128)); // Followed Hyperlink (purple)

        // Add a column chart that will inherit the theme colors
        int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 10);
        Chart chart = sheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", false);
        chart.NSeries.CategoryData = "A2:A4";

        // Optionally set a built‑in chart style (1‑48). Style 2 works well with the Office theme.
        chart.Style = 2;

        // Save the workbook
        workbook.Save("ChartWithOfficeTheme.xlsx");
    }
}
