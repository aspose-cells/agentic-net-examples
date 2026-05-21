using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

class ChangeChartThemeToOffice
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add sample data for the chart
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
        chart.NSeries.Add("B2:B4", false);
        chart.NSeries.CategoryData = "A2:A4";

        // Define the built‑in "Office" theme colors (12 colors)
        Color[] officeColors = new Color[]
        {
            Color.White,                     // Background1
            Color.Black,                     // Text1
            Color.White,                     // Background2
            Color.Black,                     // Text2
            Color.FromArgb(0, 112, 192),     // Accent1 (Blue)
            Color.FromArgb(255, 192, 0),     // Accent2 (Orange)
            Color.FromArgb(112, 48, 160),    // Accent3 (Purple)
            Color.FromArgb(0, 176, 80),      // Accent4 (Green)
            Color.FromArgb(255, 0, 0),       // Accent5 (Red)
            Color.FromArgb(255, 0, 255),     // Accent6 (Magenta)
            Color.FromArgb(0, 0, 255),       // Hyperlink (Blue)
            Color.FromArgb(128, 0, 128)      // Followed Hyperlink (Purple)
        };

        // Apply the "Office" theme to the workbook (charts inherit the workbook theme)
        workbook.CustomTheme("Office", officeColors);

        // Save the workbook with the themed chart
        workbook.Save("ChartWithOfficeTheme.xlsx");
    }
}