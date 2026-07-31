// Title: Apply a Custom Theme and Set Series Colors in a Column Chart with Aspose.Cells for .NET
// Description: Creates a workbook, adds sample data, inserts a column chart, defines a 12‑color custom theme, applies the theme, and assigns Accent1 to the first series and Accent2 to the second series using ThemeColor objects before saving the file.
// Keywords: Aspose.Cells custom theme | column chart series color | C# Excel theme color | ThemeColor Aspose.Cells | .NET chart styling | Excel workbook branding | custom theme colors Aspose
// Common Searches: how to apply a custom theme in Aspose.Cells .NET | change individual series colors in an Excel chart using Aspose.Cells | Aspose.Cells ThemeColor example C# | set accent colors for chart series Aspose.Cells | custom Excel theme with 12 colors Aspose
// Developer Intent: Apply a predefined custom theme to a workbook and map specific accent colors to each column‑chart series.
// Use Cases: Generate a branded Excel report where the workbook uses a corporate color palette and each chart series reflects the brand’s accent colors. | Reuse a single custom theme across multiple worksheets while customizing series colors for distinct data sets. | Override only the series colors of a chart while keeping the theme’s background, text, and hyperlink colors unchanged.
// AI Prompts: Write C# code with Aspose.Cells that creates a 12‑color custom theme, applies it to a workbook, and sets the first two column‑chart series to Accent1 and Accent2. | Show how to modify the border ThemeColor of a chart series after a custom theme has been applied in Aspose.Cells for .NET. | Explain how to retrieve the colors from a custom theme and reuse them for other charts in the same workbook.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// Creates a workbook, adds sample data, inserts a column chart, defines a 12‑color custom theme, applies the theme, and assigns Accent1 to the first series and Accent2 to the second series using ThemeColor objects before saving the file.
class ApplyThemeAndCustomSeriesColors
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the column chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("Jan");
        sheet.Cells["A3"].PutValue("Feb");
        sheet.Cells["A4"].PutValue("Mar");

        sheet.Cells["B1"].PutValue("Series1");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["B4"].PutValue(30);

        sheet.Cells["C1"].PutValue("Series2");
        sheet.Cells["C2"].PutValue(15);
        sheet.Cells["C3"].PutValue(25);
        sheet.Cells["C4"].PutValue(35);

        // Add a column chart to the worksheet
        int chartIdx = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 10);
        Chart chart = sheet.Charts[chartIdx];

        // Set the data range for the series and categories
        chart.NSeries.Add("B2:C4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Define a custom theme (12 colors as required by CustomTheme)
        Color[] customColors = new Color[]
        {
            Color.White,                     // Background1
            Color.Black,                     // Text1
            Color.LightGray,                 // Background2
            Color.DarkGray,                  // Text2
            Color.FromArgb(79, 129, 189),    // Accent1
            Color.FromArgb(192, 80, 77),     // Accent2
            Color.FromArgb(155, 187, 89),    // Accent3
            Color.FromArgb(128, 100, 162),   // Accent4
            Color.FromArgb(75, 172, 198),    // Accent5
            Color.FromArgb(247, 150, 70),    // Accent6
            Color.Blue,                      // Hyperlink
            Color.Purple                     // Followed Hyperlink
        };

        // Apply the custom theme to the workbook
        workbook.CustomTheme("MyCustomTheme", customColors);

        // Customize individual series colors using the theme's accent colors
        // First series -> Accent1
        chart.NSeries[0].Area.ForegroundColor = customColors[4];
        chart.NSeries[0].Border.ThemeColor = new ThemeColor(ThemeColorType.Accent1, 0);

        // Second series -> Accent2
        chart.NSeries[1].Area.ForegroundColor = customColors[5];
        chart.NSeries[1].Border.ThemeColor = new ThemeColor(ThemeColorType.Accent2, 0);

        // Save the workbook with the themed chart
        workbook.Save("ColumnChartWithCustomTheme.xlsx");
    }
}
