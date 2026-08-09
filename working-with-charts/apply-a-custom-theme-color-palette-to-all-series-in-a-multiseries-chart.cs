// Title: Apply a custom theme and monochromatic palette to every series in a multi‑series chart (C# Aspose.Cells)
// Description: Creates a workbook, defines a 12‑color custom theme, applies it with Workbook.CustomTheme, adds a column chart with two series, sets the data ranges, and changes all series colors to the first monochromatic palette using NSeries.ChangeColors, then saves the file.
// Keywords: Aspose.Cells custom theme | Workbook.CustomTheme C# | Chart NSeries.ChangeColors | monochromatic palette Aspose.Cells | .NET chart series colors | multi‑series column chart | apply theme to chart | ChartColorPaletteType | C# Excel chart styling
// Common Searches: Aspose.Cells set custom workbook theme | change colors of all chart series Aspose.Cells .NET | apply monochromatic palette to multi‑series chart | C# example for chart theming with Aspose.Cells
// Developer Intent: The developer wants to programmatically apply a 12‑color custom workbook theme and a single monochromatic color palette to every series in a multi‑series chart using Aspose.Cells for .NET.
// Use Cases: Generate Excel reports where every chart automatically follows a corporate 12‑color theme. | Create a column chart with multiple series that shares a consistent monochromatic accent without setting each series color individually. | Update existing workbooks to adopt a new theme and unified series palette in a single API call.
// AI Prompts: Write C# code with Aspose.Cells to apply a custom 12‑color theme and then set ChartColorPaletteType.MonochromaticPalette2 for all series in a line chart. | Explain how Workbook.CustomTheme and NSeries.ChangeColors interact to style charts in Aspose.Cells. | Provide step‑by‑step instructions to build a multi‑series chart, assign category and value ranges, and apply a custom theme plus a monochromatic palette.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsCustomThemeDemo
{
    // Creates a workbook, defines a 12‑color custom theme, applies it with Workbook.CustomTheme, adds a column chart with two series, sets the data ranges, and changes all series colors to the first monochromatic palette using NSeries.ChangeColors, then saves the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for a multi‑series chart
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

            // Define a custom theme (12 colors as required)
            Color[] customColors = new Color[]
            {
                Color.FromArgb(255, 255, 255), // Background1
                Color.FromArgb(0, 0, 0),       // Text1
                Color.FromArgb(240, 240, 240), // Background2
                Color.FromArgb(50, 50, 50),    // Text2
                Color.FromArgb(0, 112, 192),   // Accent1
                Color.FromArgb(255, 192, 0),   // Accent2
                Color.FromArgb(112, 173, 71),  // Accent3
                Color.FromArgb(255, 0, 0),     // Accent4
                Color.FromArgb(0, 176, 80),    // Accent5
                Color.FromArgb(0, 176, 240),   // Accent6
                Color.FromArgb(0, 0, 255),     // Hyperlink
                Color.FromArgb(128, 0, 128)    // Followed Hyperlink
            };

            // Apply the custom theme to the workbook
            workbook.CustomTheme("MyCustomTheme", customColors);

            // Add a column chart that contains both series
            int chartIdx = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 12);
            Chart chart = sheet.Charts[chartIdx];

            // Set the data range for the series (both Series1 and Series2)
            chart.NSeries.Add("B1:C4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Apply a monochromatic palette to all series in the collection
            // Here we use the first monochromatic palette (Accent1 gradient)
            chart.NSeries.ChangeColors(ChartColorPaletteType.MonochromaticPalette1);

            // Save the workbook
            workbook.Save("CustomTheme_MultiSeriesChart.xlsx");
        }
    }
}
