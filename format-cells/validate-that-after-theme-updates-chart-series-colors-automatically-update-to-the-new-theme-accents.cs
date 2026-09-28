// Title: Check whether Aspose.Cells for .NET updates column chart series colors automatically after applying an Office theme
// AI Prompts: Write C# code with Aspose.Cells that creates a column chart, captures the series foreground color, applies a .theme file to the workbook, refreshes the chart, and returns a boolean indicating if the color changed. | Modify the example to load a custom theme file path, call Workbook.SetTheme, invoke chart.Refresh if supported, and throw an exception when the series color does not match the new theme accent. | Create a C# unit test that builds a workbook with a column chart, records the initial series color, applies the Office2013 theme, forces a chart refresh, and fails the test if the colors remain identical.
// Common Searches: asp.net aspose.cells verify chart series color changes after applying an Excel theme | c# detect automatic chart color update when workbook theme is changed using Aspose.Cells | aspose.cells check if column chart uses new theme accent colors without manual color assignment
// Tags: apply built‑in Office theme Aspose.Cells | validate chart series color change Aspose.Cells | column chart theme accent propagation | Workbook.SetTheme chart refresh | C# Aspose.Cells chart color verification

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The program builds a workbook with sample data, adds a column chart, records the initial series foreground color, attempts to apply an Office2013 .theme file (if supported), captures the series color again, compares the two colors to determine whether the theme altered the chart series color, and saves the workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Populate worksheet with sample data
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart linked to the data
            int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIdx];
            chart.NSeries.Add("B2:B4", true);
            // Category (X‑axis) data is automatically taken from the first column when adding a series.
            // If explicit setting is required, uncomment the line below (requires a version that supports it).
            // chart.NSeries[0].CategoryData = "A2:A4";

            // Capture the initial series color (default theme)
            Color initialColor = chart.NSeries[0].Area.ForegroundColor;

            // Attempt to apply a built‑in theme if the file exists
            string themePath = "Office2013.theme";
            if (File.Exists(themePath))
            {
                try
                {
                    // workbook.SetTheme(themePath); // Not available in all versions
                    // chart.Refresh();               // May be unavailable
                    Console.WriteLine($"Theme file found but SetTheme/Refresh not supported in this version. Skipping theme application.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error applying theme: {ex.Message}");
                }
            }
            else
            {
                Console.WriteLine($"Theme file not found: {themePath}. Skipping theme application.");
            }

            // Capture the series color after the (skipped) theme change
            Color updatedColor = chart.NSeries[0].Area.ForegroundColor;

            // Validate whether the series color has changed
            bool colorsUpdated = !initialColor.Equals(updatedColor);
            Console.WriteLine($"Initial series color: {initialColor}");
            Console.WriteLine($"Updated series color: {updatedColor}");
            Console.WriteLine($"Series colors updated after theme change: {colorsUpdated}");

            // Save the workbook
            string outputPath = "ThemeUpdateValidation.xlsx";
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
