// Title: Apply a Corporate Custom Theme to a Waterfall Chart with Aspose.Cells for .NET (C#)
// Description: Creates a new workbook, adds sample waterfall data, inserts a Waterfall chart, defines its data range, builds a 12‑color corporate palette, applies it via workbook.CustomTheme, and saves the Excel file with the branded chart.
// Keywords: Aspose.Cells | C# | Waterfall chart | custom theme | corporate branding | Excel workbook | CustomTheme API | chart colors | Excel automation | sample code
// Common Searches: Aspose.Cells apply custom theme to chart | C# set corporate colors for Excel Waterfall chart | How to use CustomTheme with Aspose.Cells .NET | Branding Excel charts programmatically | Waterfall chart theme example Aspose.Cells
// Developer Intent: Apply a predefined corporate color palette to a Waterfall chart in an Excel workbook using Aspose.Cells for .NET.
// Use Cases: Generate financial waterfall reports that automatically follow company branding. | Create a batch of Excel workbooks with a consistent corporate theme for all embedded charts. | Automate production of presentation‑ready charts that use the organization’s official color set.
// AI Prompts: Show C# code to apply the same corporate CustomTheme to a Column chart with Aspose.Cells. | Demonstrate loading theme colors from a JSON file and applying them via workbook.CustomTheme. | Explain how to modify an existing theme after a chart has been created without rebuilding the workbook. | Provide a GitHub‑style README snippet describing this example for the Aspose.Cells repository.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a new workbook, adds sample waterfall data, inserts a Waterfall chart, defines its data range, builds a 12‑color corporate palette, applies it via workbook.CustomTheme, and saves the Excel file with the branded chart.
class ApplyCorporateThemeToWaterfallChart
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for a Waterfall chart
        // Column A: Categories, Column B: Values
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("Start");
        sheet.Cells["A3"].PutValue("Increase");
        sheet.Cells["A4"].PutValue("Decrease");
        sheet.Cells["A5"].PutValue("End");

        sheet.Cells["B1"].PutValue("Amount");
        sheet.Cells["B2"].PutValue(5000);
        sheet.Cells["B3"].PutValue(2000);
        sheet.Cells["B4"].PutValue(-1500);
        sheet.Cells["B5"].PutValue(5500);

        // Add a Waterfall chart
        int chartIndex = sheet.Charts.Add(ChartType.Waterfall, 7, 0, 25, 10);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data range for the chart (including headers)
        chart.SetChartDataRange("A1:B5", true);
        // Define the category (X) axis data
        chart.NSeries.CategoryData = "A2:A5";

        // -------------------------------------------------
        // Apply a corporate custom theme (12 colors required)
        // -------------------------------------------------
        Color[] corporateColors = new Color[]
        {
            Color.FromArgb(255, 255, 255), // Background1 (white)
            Color.FromArgb(0, 0, 0),       // Text1 (black)
            Color.FromArgb(240, 240, 240), // Background2 (light gray)
            Color.FromArgb(80, 80, 80),    // Text2 (dark gray)
            Color.FromArgb(0, 112, 192),   // Accent1 (corporate blue)
            Color.FromArgb(255, 192, 0),   // Accent2 (corporate orange)
            Color.FromArgb(112, 173, 71),  // Accent3 (corporate green)
            Color.FromArgb(191, 0, 0),     // Accent4 (corporate red)
            Color.FromArgb(255, 0, 255),   // Accent5 (magenta)
            Color.FromArgb(0, 176, 80),    // Accent6 (secondary green)
            Color.FromArgb(0, 0, 255),     // Hyperlink (blue)
            Color.FromArgb(128, 0, 128)    // Followed Hyperlink (purple)
        };

        // Apply the custom theme to the workbook
        workbook.CustomTheme("CorporateBranding", corporateColors);

        // Save the workbook with the themed Waterfall chart
        workbook.Save("WaterfallChartWithCorporateTheme.xlsx");
    }
}
