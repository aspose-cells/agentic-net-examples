// Title: C# – Replace Workbook Theme Accent4 with Corporate Green and Verify Chart Colors using Aspose.Cells
// Description: Loads a workbook, changes the theme’s Accent4 color to a corporate green (RGB 0,128,0) via SetThemeColor, scans all worksheets and charts to confirm each series uses the updated Accent4 fill, logs the results, and saves the modified file.
// Keywords: Aspose.Cells C# set theme color | replace Accent4 theme color | corporate green workbook | chart series color verification | SetThemeColor Aspose | theme color validation .NET | Excel theme customization Aspose.Cells | chart fill format check | theme color Accent4 | Excel branding automation
// Common Searches: Aspose.Cells change theme accent color | How to set Accent4 to custom RGB in C# | Validate chart series theme color after workbook theme change | C# code to update Excel theme colors with Aspose | Check if chart series uses theme accent in Aspose.Cells
// Developer Intent: Update the workbook’s Accent4 theme color to a corporate green and ensure all chart series reflect the new color.
// Use Cases: Standardize corporate branding across Excel reports by programmatically updating theme colors. | Automated QA to confirm chart colors match brand guidelines after theme modifications. | Create versioned workbooks for different brand palettes without manual editing. | Integrate theme updates into a CI pipeline for report generation.
// AI Prompts: Generate C# Aspose.Cells code that changes the Accent4 theme color to #008000 and returns a list of chart series that still use the previous color. | Write a method that iterates through all charts in a workbook and verifies each series' fill references Accent4, outputting mismatches. | Explain how to revert an Accent4 theme color change and re‑run the verification using Aspose.Cells.

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// Loads a workbook, changes the theme’s Accent4 color to a corporate green (RGB 0,128,0) via SetThemeColor, scans all worksheets and charts to confirm each series uses the updated Accent4 fill, logs the results, and saves the modified file.
class ThemeReplaceDemo
{
    static void Main()
    {
        const string inputPath = "InputWithChart.xlsx";
        const string outputPath = "OutputWithCorporateGreen.xlsx";

        try
        {
            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook that contains charts
            Workbook workbook = new Workbook(inputPath);

            // Define the corporate green color (example: RGB 0,128,0)
            Color corporateGreen = Color.FromArgb(0, 128, 0);

            // Replace the theme's Accent4 color with the corporate green
            workbook.SetThemeColor(ThemeColorType.Accent4, corporateGreen);

            // Verify that chart series use the updated Accent4 theme color
            foreach (Worksheet ws in workbook.Worksheets)
            {
                foreach (Chart chart in ws.Charts)
                {
                    for (int i = 0; i < chart.NSeries.Count; i++)
                    {
                        var series = chart.NSeries[i];
                        var fill = series.Area.FillFormat;

                        // Check only solid fill formats
                        if (fill != null && fill.Type == FillType.Solid)
                        {
                            CellsColor cellsColor = fill.SolidFill.CellsColor;
                            ThemeColor themeColor = cellsColor?.ThemeColor;

                            if (themeColor != null && themeColor.ColorType == ThemeColorType.Accent4)
                            {
                                Console.WriteLine($"Chart '{chart.Name}' series {i} correctly uses Accent4.");
                            }
                            else
                            {
                                Console.WriteLine($"Chart '{chart.Name}' series {i} does NOT use Accent4.");
                            }
                        }
                    }
                }
            }

            // Save the workbook with the updated theme
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
