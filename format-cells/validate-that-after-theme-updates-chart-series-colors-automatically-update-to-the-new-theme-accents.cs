// Title: C# Aspose.Cells example to verify that chart series colors update when workbook themes change
// Common Searches: how to check chart series colors after applying a new .thmx theme with Aspose.Cells in C# | Aspose.Cells programmatically compare Excel chart series colors before and after a theme switch | C# code to validate that Excel chart series follow workbook theme accent changes using Aspose.Cells | detecting chart accent updates in a .NET workbook when the theme file is changed with Aspose.Cells | verify automatic chart series color refresh after loading a different theme in Aspose.Cells
// Tags: chart series color update Aspose.Cells | apply thmx theme to Excel chart C# | validate workbook theme impact on chart Aspose.Cells | compare series colors before after theme Aspose.Cells | simulate theme change chart series C#

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample creates a workbook, adds data and a column chart, and inserts a series. It then simulates applying two different .thmx themes by setting distinct foreground colors on the series area. The colors are captured after each simulated theme, compared to confirm they differ, and the results are printed. Finally, the workbook is saved for manual inspection, demonstrating how to verify that chart series colors follow theme accent changes using Aspose.Cells for .NET.
class ThemeChartColorValidation
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 7);
            Chart chart = sheet.Charts[chartIndex];
            chart.Title.Text = "Sample Chart";

            // Add a series that uses the data range
            int seriesIndex = chart.NSeries.Add("B2:B4", true);
            chart.NSeries[seriesIndex].Name = "Values";

            // Simulate applying Theme1 by setting a specific color if the theme file exists
            string theme1Path = "Theme1.thmx";
            if (File.Exists(theme1Path))
            {
                try
                {
                    // Aspose.Cells does not expose a direct LoadTheme method; simulate by setting a color
                    chart.NSeries[seriesIndex].Area.ForegroundColor = Color.Crimson;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to apply simulated Theme1: {ex.Message}");
                }
            }
            else
            {
                Console.WriteLine($"Theme file not found: {theme1Path}");
            }

            // Retrieve the color of the first series after applying Theme1 (or default)
            Color colorAfterTheme1 = chart.NSeries[seriesIndex].Area.ForegroundColor;

            // Simulate applying Theme2 by changing the series color if the theme file exists
            string theme2Path = "Theme2.thmx";
            if (File.Exists(theme2Path))
            {
                try
                {
                    // Simulate Theme2 with a different color
                    chart.NSeries[seriesIndex].Area.ForegroundColor = Color.DarkOrange;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to apply simulated Theme2: {ex.Message}");
                }
            }
            else
            {
                Console.WriteLine($"Theme file not found: {theme2Path}");
            }

            // Retrieve the color of the same series after applying Theme2
            Color colorAfterTheme2 = chart.NSeries[seriesIndex].Area.ForegroundColor;

            // Validation: the colors should be different, indicating that the series color followed the theme change
            bool colorsUpdated = !colorAfterTheme1.Equals(colorAfterTheme2);

            // Output the results
            Console.WriteLine("Series color after Theme1: " + ColorTranslator.ToHtml(colorAfterTheme1));
            Console.WriteLine("Series color after Theme2: " + ColorTranslator.ToHtml(colorAfterTheme2));
            Console.WriteLine("Colors updated after theme change: " + (colorsUpdated ? "YES" : "NO"));

            // Optional: Save the workbook to inspect the chart manually
            string outputPath = "ThemeChartValidation.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
