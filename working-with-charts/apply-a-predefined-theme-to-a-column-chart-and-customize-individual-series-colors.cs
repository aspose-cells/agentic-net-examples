using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

class ApplyThemeAndCustomSeriesColors
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the chart
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

        // Add a column chart
        int chartIdx = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 10);
        Chart chart = sheet.Charts[chartIdx];

        // Set the data range for the series and categories
        chart.NSeries.Add("B2:C4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Define a custom theme (12 colors as required by CustomTheme)
        Color[] customColors = new Color[]
        {
            Color.FromArgb(255, 255, 255), // Background1
            Color.FromArgb(0, 0, 0),       // Text1
            Color.FromArgb(240, 240, 240), // Background2
            Color.FromArgb(80, 80, 80),    // Text2
            Color.FromArgb(79, 129, 189),  // Accent1
            Color.FromArgb(192, 80, 77),   // Accent2
            Color.FromArgb(155, 187, 89),  // Accent3
            Color.FromArgb(128, 100, 162), // Accent4
            Color.FromArgb(75, 172, 198),  // Accent5
            Color.FromArgb(247, 150, 70),  // Accent6
            Color.FromArgb(0, 0, 255),     // Hyperlink
            Color.FromArgb(128, 0, 128)    // Followed Hyperlink
        };

        // Apply the custom theme to the workbook
        workbook.CustomTheme("MyCustomTheme", customColors);

        // Customize individual series colors using the theme's accent colors
        chart.NSeries[0].Area.ForegroundColor = customColors[4]; // Accent1 for Series1
        chart.NSeries[1].Area.ForegroundColor = customColors[5]; // Accent2 for Series2

        // Save the workbook with the chart
        workbook.Save("ColumnChartWithCustomTheme.xlsx");
    }
}