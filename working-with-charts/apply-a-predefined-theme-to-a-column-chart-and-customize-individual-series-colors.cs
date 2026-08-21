// Title: C# – Apply a Custom Theme and Set Series Colors in an Aspose.Cells Column Chart
// Description: Creates a new workbook, defines a 12‑color custom theme with Workbook.CustomTheme, adds a column chart, and assigns Accent1 and Accent2 from the theme to the first two series before saving the file as an Excel workbook.
// Keywords: Aspose.Cells | C# | .NET | custom workbook theme | column chart | chart series colors | Workbook.CustomTheme | Chart.NSeries | Excel chart styling | Color.FromArgb | GitHub example
// Common Searches: Aspose.Cells apply custom theme to workbook C# | change individual series colors in Aspose.Cells column chart | C# example for Workbook.CustomTheme and chart series colors
// Developer Intent: Generate an Excel file, apply a 12‑color custom theme, insert a column chart, and color each series using theme accent colors.
// Use Cases: Corporate sales reports that must follow brand colors across chart series. | Automated dashboard generation where each product line uses a distinct theme accent. | Presentation‑ready Excel files with consistent styling for multiple chart series.
// AI Prompts: Write C# code with Aspose.Cells to create a workbook, define a 12‑color custom theme, add a column chart, and set series 0 to Accent1 and series 1 to Accent2. | Explain how Workbook.CustomTheme works in Aspose.Cells and how to reference its colors when customizing chart series. | Provide step‑by‑step instructions to change individual series colors in an Aspose.Cells column chart after applying a custom theme, including saving the workbook.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// Creates a new workbook, defines a 12‑color custom theme with Workbook.CustomTheme, adds a column chart, and assigns Accent1 and Accent2 from the theme to the first two series before saving the file as an Excel workbook.
class ApplyThemeAndCustomizeSeries
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the column chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("Q1");
        sheet.Cells["A3"].PutValue("Q2");
        sheet.Cells["A4"].PutValue("Q3");

        sheet.Cells["B1"].PutValue("Series1");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["B4"].PutValue(30);

        sheet.Cells["C1"].PutValue("Series2");
        sheet.Cells["C2"].PutValue(15);
        sheet.Cells["C3"].PutValue(25);
        sheet.Cells["C4"].PutValue(35);

        // Define a custom theme (12 colors as required)
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

        // Add a column chart
        int chartIdx = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 10);
        Chart chart = sheet.Charts[chartIdx];

        // Set the data range for the chart
        chart.NSeries.Add("B2:C4", true);   // Series values
        chart.NSeries.CategoryData = "A2:A4"; // Categories

        // Customize individual series colors
        // Series 0 (Series1) -> use Accent1 from the theme
        chart.NSeries[0].Area.ForegroundColor = customColors[4]; // Accent1
        // Series 1 (Series2) -> use Accent2 from the theme
        chart.NSeries[1].Area.ForegroundColor = customColors[5]; // Accent2

        // Optionally, change the overall palette (demonstrates ChangeColors usage)
        // Here we keep the default palette, so this line is commented out.
        // chart.NSeries.ChangeColors(ChartColorPaletteType.Monochrome);

        // Save the workbook
        workbook.Save("ColumnChartWithCustomTheme.xlsx");
    }
}
