// Title: Replace a workbook's Theme Accent4 with a corporate green and validate chart series colors using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# replace workbook theme Accent4 with custom green color | how to check chart series colors after changing theme accent in Aspose.Cells | C# code to set theme accent color and verify Excel chart colors with Aspose.Cells | update Excel theme accent to corporate green and validate charts using Aspose.Cells .NET | verify that chart series use the new Accent4 color after theme modification in C#
// Tags: set workbook theme accent color Aspose.Cells | replace theme Accent4 with custom color C# | validate chart series theme color Aspose.Cells | apply corporate green theme to Excel workbook | chart color verification after theme change Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example loads an existing Excel workbook, defines a corporate green (RGB 0,128,0), replaces the workbook's ThemeColorType.Accent4 with this color via SetThemeColor, iterates through all worksheets and charts to ensure each series' fill color matches the new accent, logs any mismatches, and saves the updated workbook. Exception handling is included for missing files and runtime errors.
class ThemeAccentUpdater
{
    static void Main()
    {
        try
        {
            // Input workbook path
            string inputPath = "input.xlsx";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file '{inputPath}' not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Define the corporate green color (e.g., #008000)
            Color corporateGreen = Color.FromArgb(0, 128, 0);

            // Replace the theme's Accent4 color with the corporate green
            workbook.SetThemeColor(ThemeColorType.Accent4, corporateGreen);

            // Verify that chart series use the updated Accent4 color
            VerifyChartColors(workbook, corporateGreen);

            // Save the modified workbook
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    /// <param name="workbook">The workbook to inspect.</param>
    /// <param name="expectedColor">The color that should be used for series.</param>
    private static void VerifyChartColors(Workbook workbook, Color expectedColor)
    {
        try
        {
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all charts on the worksheet
                foreach (Chart chart in sheet.Charts)
                {
                    // Iterate through each series in the chart
                    foreach (Series series in chart.NSeries)
                    {
                        // The series area foreground color reflects the fill color of the series
                        Color seriesColor = series.Area.ForegroundColor;

                        // Compare with the expected Accent4 color
                        if (seriesColor.ToArgb() != expectedColor.ToArgb())
                        {
                            Console.WriteLine(
                                $"Warning: Chart '{chart.Name}' on sheet '{sheet.Name}' has a series with color {seriesColor} " +
                                $"which does not match the updated Accent4 color {expectedColor}.");
                        }
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error while verifying chart colors: {ex.Message}");
        }
    }
}
