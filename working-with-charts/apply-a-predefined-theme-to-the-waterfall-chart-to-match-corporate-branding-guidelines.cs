using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsThemeWaterfallDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // -------------------------------------------------
            // Prepare sample data for a Waterfall chart
            // -------------------------------------------------
            // Header row
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");

            // Data rows
            sheet.Cells["A2"].PutValue("Start");
            sheet.Cells["B2"].PutValue(5000);

            sheet.Cells["A3"].PutValue("Revenue");
            sheet.Cells["B3"].PutValue(8000);

            sheet.Cells["A4"].PutValue("Cost");
            sheet.Cells["B4"].PutValue(-3000);

            sheet.Cells["A5"].PutValue("Profit");
            sheet.Cells["B5"].PutValue(2000);

            // -------------------------------------------------
            // Add a Waterfall chart
            // -------------------------------------------------
            int chartIndex = sheet.Charts.Add(ChartType.Waterfall, 7, 0, 25, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart (including headers)
            chart.SetChartDataRange("A1:B5", true);

            // -------------------------------------------------
            // Define a corporate custom theme (12 colors)
            // -------------------------------------------------
            Color[] corporateColors = new Color[]
            {
                Color.FromArgb(255, 255, 255), // Background1 (white)
                Color.FromArgb(0, 0, 0),       // Text1 (black)
                Color.FromArgb(240, 240, 240), // Background2 (light gray)
                Color.FromArgb(80, 80, 80),    // Text2 (dark gray)
                Color.FromArgb(0, 112, 192),   // Accent1 (corporate blue)
                Color.FromArgb(255, 192, 0),   // Accent2 (corporate amber)
                Color.FromArgb(112, 173, 71),  // Accent3 (corporate green)
                Color.FromArgb(192, 0, 0),     // Accent4 (corporate red)
                Color.FromArgb(255, 0, 255),   // Accent5 (magenta)
                Color.FromArgb(0, 176, 80),    // Accent6 (secondary green)
                Color.FromArgb(0, 0, 255),     // Hyperlink (blue)
                Color.FromArgb(128, 0, 128)    // Followed Hyperlink (purple)
            };

            // Apply the custom theme to the workbook
            workbook.CustomTheme("CorporateTheme", corporateColors);

            // -------------------------------------------------
            // Save the workbook
            // -------------------------------------------------
            workbook.Save("WaterfallWithCorporateTheme.xlsx");
        }
    }
}