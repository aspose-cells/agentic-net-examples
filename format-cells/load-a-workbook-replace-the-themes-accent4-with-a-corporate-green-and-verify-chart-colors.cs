// Title: C# – Change Excel Theme Accent4 to Corporate Green and Verify Chart Colors with Aspose.Cells
// Description: Loads an Excel workbook, replaces the theme's Accent4 color with a corporate green (RGB 0,128,0), applies the updated theme to chart series, prints verification details for each chart, and saves the workbook.
// Keywords: Aspose.Cells C# set theme color | Excel Accent4 change programmatically | apply corporate green to chart series | verify chart theme color Aspose.Cells | Workbook.SetThemeColor example | ThemeColorType Accent4 | update Excel theme with Aspose | chart series color verification | .NET Excel theme customization
// Common Searches: how to change Accent4 theme color in Excel using Aspose.Cells C# | replace Excel theme color with corporate branding programmatically | verify chart series color after theme update Aspose.Cells | set custom RGB color for Excel theme Accent4 .NET | iterate charts and read applied theme color Aspose.Cells
// Developer Intent: Replace the workbook's Accent4 theme color with a corporate green and confirm that chart series reflect the new color.
// Use Cases: Enforce corporate branding by updating the Accent4 theme across all worksheets and charts. | Automate quality checks that ensure chart series use the intended brand color before publishing reports. | Log theme‑color details for each chart to aid troubleshooting of visual inconsistencies.
// AI Prompts: Generate C# code using Aspose.Cells to set the Accent4 theme color to a specific RGB value and apply it to every chart series in a workbook. | Create a method that iterates through all charts in an Excel file, prints the ThemeColor type, tint, and verification status for each series. | Explain how to revert an Accent4 theme color change back to its original value with Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// Loads an Excel workbook, replaces the theme's Accent4 color with a corporate green (RGB 0,128,0), applies the updated theme to chart series, prints verification details for each chart, and saves the workbook.
class Program
{
    static void Main()
    {
        // Load the existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Define the corporate green color
        Color corporateGreen = Color.FromArgb(0, 128, 0); // Dark green

        // Replace the theme's Accent4 color with the corporate green
        workbook.SetThemeColor(ThemeColorType.Accent4, corporateGreen);

        // Iterate through all worksheets and their charts
        foreach (Worksheet ws in workbook.Worksheets)
        {
            foreach (Chart chart in ws.Charts)
            {
                // Ensure the chart has at least one series
                if (chart.NSeries.Count > 0)
                {
                    // Apply the Accent4 theme color to the first series (for demonstration)
                    // This makes the series use the updated theme color
                    chart.NSeries[0].Area.FillFormat.SolidFill.CellsColor.ThemeColor =
                        new ThemeColor(ThemeColorType.Accent4, 0);

                    // Retrieve the theme color applied to the series
                    CellsColor seriesColor = chart.NSeries[0].Area.FillFormat.SolidFill.CellsColor;
                    ThemeColor appliedTheme = seriesColor.ThemeColor;

                    // Output verification information
                    Console.WriteLine($"Worksheet: {ws.Name}, Chart: {chart.Name}");
                    Console.WriteLine($"Applied Theme Color Type: {appliedTheme.ColorType}");
                    Console.WriteLine($"Is Accent4: {appliedTheme.ColorType == ThemeColorType.Accent4}");
                    Console.WriteLine($"Tint: {appliedTheme.Tint}");
                    Console.WriteLine();
                }
            }
        }

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}
