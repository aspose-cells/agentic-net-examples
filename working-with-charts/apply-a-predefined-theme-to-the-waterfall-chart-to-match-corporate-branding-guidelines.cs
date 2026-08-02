// Title: C# – Apply a Corporate Custom Theme to a Waterfall Chart using Aspose.Cells
// Description: Demonstrates how to create a workbook, add sample financial data, generate a Waterfall chart, define a 12‑color corporate palette, apply it as a custom theme with Aspose.Cells for .NET, and save the Excel file with branding‑consistent chart colors.
// Keywords: Aspose.Cells | C# | .NET | Waterfall chart | custom theme | corporate colors | Excel branding | CustomTheme method | chart styling | financial waterfall | example code
// Common Searches: Aspose.Cells apply custom theme to chart | C# Waterfall chart corporate branding | How to set Excel theme colors with Aspose.Cells | CustomTheme example for Waterfall chart .NET | Apply corporate palette to Excel chart programmatically
// Developer Intent: Add a predefined corporate color palette to a Waterfall chart and save the themed workbook.
// Use Cases: Produce quarterly financial waterfall charts that match company brand guidelines. | Create multiple Excel charts in one workbook that share a unified corporate theme for presentations. | Automate generation of branded Excel reports for distribution to stakeholders.
// AI Prompts: Show how to change the corporate theme colors after the chart is created. | Provide code to reuse the same custom theme for column, line, and pie charts with Aspose.Cells. | Explain how to retrieve, edit, or replace an existing custom theme in a workbook.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Demonstrates how to create a workbook, add sample financial data, generate a Waterfall chart, define a 12‑color corporate palette, apply it as a custom theme with Aspose.Cells for .NET, and save the Excel file with branding‑consistent chart colors.
class ApplyCorporateThemeToWaterfallChart
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // -------------------------------------------------
        // 1. Prepare sample data for the Waterfall chart
        // -------------------------------------------------
        // Header row
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Value");

        // Sample categories and values
        string[] categories = { "Start", "Revenue", "Cost", "Profit", "End" };
        double[] values = { 0, 120, -40, 80, 0 };

        for (int i = 0; i < categories.Length; i++)
        {
            sheet.Cells[i + 2, 0].PutValue(categories[i]);   // Column A
            sheet.Cells[i + 2, 1].PutValue(values[i]);      // Column B
        }

        // -------------------------------------------------
        // 2. Add a Waterfall chart and bind the data range
        // -------------------------------------------------
        int chartIndex = sheet.Charts.Add(ChartType.Waterfall, 7, 0, 25, 10);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data range for the series (values) and categories
        chart.NSeries.Add("B2:B6", false);
        chart.NSeries.CategoryData = "A2:A6";

        // Optional: give the series a name
        chart.NSeries[0].Name = "Financial Flow";

        // -------------------------------------------------
        // 3. Define a corporate custom theme (12 colors)
        // -------------------------------------------------
        // The order of colors follows the ThemeColorType index:
        // 0-Background1, 1-Text1, 2-Background2, 3-Text2,
        // 4-Accent1, 5-Accent2, 6-Accent3, 7-Accent4,
        // 8-Accent5, 9-Accent6, 10-Hyperlink, 11-FollowedHyperlink
        Color[] corporateColors = new Color[]
        {
            Color.FromArgb(255, 255, 255), // Background1 – White
            Color.FromArgb(0, 0, 0),       // Text1 – Black
            Color.FromArgb(240, 240, 240), // Background2 – Light Gray
            Color.FromArgb(80, 80, 80),    // Text2 – Dark Gray
            Color.FromArgb(0, 112, 192),   // Accent1 – Corporate Blue
            Color.FromArgb(255, 192, 0),   // Accent2 – Corporate Gold
            Color.FromArgb(0, 176, 80),    // Accent3 – Corporate Green
            Color.FromArgb(192, 0, 0),     // Accent4 – Corporate Red
            Color.FromArgb(112, 48, 160),  // Accent5 – Corporate Purple
            Color.FromArgb(255, 0, 255),   // Accent6 – Corporate Magenta
            Color.FromArgb(0, 0, 255),     // Hyperlink – Blue
            Color.FromArgb(128, 0, 128)    // FollowedHyperlink – Purple
        };

        // Apply the custom theme to the workbook
        workbook.CustomTheme("CorporateTheme", corporateColors);

        // -------------------------------------------------
        // 4. Save the workbook with the themed Waterfall chart
        // -------------------------------------------------
        workbook.Save("WaterfallChart_CorporateTheme.xlsx");
    }
}
