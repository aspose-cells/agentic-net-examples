// Title: C# – Apply a Custom Corporate Theme to a Waterfall Chart with Aspose.Cells
// Description: Sample code that creates a workbook, adds sample waterfall data, inserts a Waterfall chart, defines a 12‑color corporate palette, applies it as a custom theme named "CorporateBranding", and saves the file. Demonstrates how to enforce branding guidelines on Excel charts using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | .NET | Waterfall chart | custom theme | corporate branding | Excel theme colors | chart palette | sample code | GitHub example | Excel automation
// Common Searches: apply custom theme to waterfall chart Aspose.Cells | Aspose.Cells corporate colors for Excel charts | C# set 12‑color theme in Aspose.Cells | how to brand Excel charts with Aspose.Cells | sample code for custom Excel theme .NET
// Developer Intent: Apply a predefined corporate theme to a Waterfall chart so the visual output follows branding guidelines.
// Use Cases: Generate quarterly financial waterfall reports that automatically use the company’s brand colors. | Create a reusable workbook template where every chart inherits the corporate theme, ensuring consistent visual identity across departments. | Automate production of multiple Excel files for marketing presentations, each containing a themed Waterfall chart.
// AI Prompts: Show how to replace the corporate palette with a different set of 12 branding colors. | Explain how to export the "CorporateBranding" theme and import it into another workbook using Aspose.Cells. | Provide code to apply the same custom theme to multiple charts on the same worksheet.

using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Sample code that creates a workbook, adds sample waterfall data, inserts a Waterfall chart, defines a 12‑color corporate palette, applies it as a custom theme named "CorporateBranding", and saves the file. Demonstrates how to enforce branding guidelines on Excel charts using Aspose.Cells for .NET.
class ApplyWaterfallTheme
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for the Waterfall chart
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["B1"].PutValue("Value");
        worksheet.Cells["A2"].PutValue("Start");
        worksheet.Cells["B2"].PutValue(100);
        worksheet.Cells["A3"].PutValue("Increase");
        worksheet.Cells["B3"].PutValue(30);
        worksheet.Cells["A4"].PutValue("Decrease");
        worksheet.Cells["B4"].PutValue(-20);
        worksheet.Cells["A5"].PutValue("End");
        worksheet.Cells["B5"].PutValue(110);

        // Add a Waterfall chart to the worksheet
        int chartIndex = worksheet.Charts.Add(ChartType.Waterfall, 7, 0, 25, 10);
        Chart chart = worksheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B5", false);          // Values
        chart.NSeries.CategoryData = "A2:A5";       // Categories

        // Define corporate branding colors (must contain 12 colors)
        Color[] corporateColors = new Color[]
        {
            Color.FromArgb(255, 255, 255), // Background1
            Color.FromArgb(0, 0, 0),       // Text1
            Color.FromArgb(240, 240, 240), // Background2
            Color.FromArgb(80, 80, 80),    // Text2
            Color.FromArgb(0, 112, 192),   // Accent1 (corporate blue)
            Color.FromArgb(255, 192, 0),   // Accent2 (corporate orange)
            Color.FromArgb(112, 173, 71),  // Accent3 (corporate green)
            Color.FromArgb(255, 0, 0),     // Accent4 (corporate red)
            Color.FromArgb(255, 255, 0),   // Accent5 (corporate yellow)
            Color.FromArgb(0, 176, 80),    // Accent6 (corporate teal)
            Color.FromArgb(0, 0, 255),     // Hyperlink
            Color.FromArgb(128, 0, 128)    // FollowedHyperlink
        };

        // Apply the custom theme to the workbook
        workbook.CustomTheme("CorporateBranding", corporateColors);

        // Save the workbook with the themed Waterfall chart
        workbook.Save("WaterfallCorporateTheme.xlsx");
    }
}
