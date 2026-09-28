// Title: Change chart accent colors by modifying the workbook Theme.ColorScheme via reflection in Aspose.Cells C#
// AI Prompts: Write C# code that uses reflection to set the Accent1 and Accent2 colors in a workbook’s Theme.ColorScheme for an Aspose.Cells chart. | Show how to change a chart’s colors when the Aspose.Cells Theme API is not directly accessible. | Provide a full example that creates a column chart, updates its theme accents, and saves the workbook.
// Common Searches: asp.net aspose.cells change chart theme accent colors using reflection | c# set workbook theme color scheme when Theme property is hidden aspose.cells | how to modify Excel chart colors by editing Theme.ColorScheme in Aspose.Cells | example updating Accent1 Accent2 colors in Aspose.Cells chart programmatically
// Tags: Aspose.Cells modify theme colors via reflection | C# update chart accent colors Aspose.Cells | set workbook Theme.ColorScheme Aspose.Cells | customize Excel chart theme colors C# | handle missing Theme API Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExample
{
    // The example creates a workbook, adds a column chart, then uses reflection to locate the hidden Theme and its ColorScheme, setting Accent1 to red and Accent2 to blue before saving the file as ModifiedThemeChart.xlsx, with fallback handling if the Theme API is unavailable.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Get the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Add a column chart to the worksheet (from row 5, column 0 to row 15, column 5)
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
                Chart chart = sheet.Charts[chartIndex];

                // Attempt to modify the workbook's theme colors using reflection (API may be unavailable)
                try
                {
                    var themeProp = workbook.GetType().GetProperty("Theme");
                    if (themeProp != null)
                    {
                        var theme = themeProp.GetValue(workbook);
                        var colorSchemeProp = theme?.GetType().GetProperty("ColorScheme");
                        var colorScheme = colorSchemeProp?.GetValue(theme);
                        if (colorScheme != null)
                        {
                            var accent1Prop = colorScheme.GetType().GetProperty("Accent1");
                            var accent2Prop = colorScheme.GetType().GetProperty("Accent2");
                            if (accent1Prop != null && accent2Prop != null)
                            {
                                accent1Prop.SetValue(colorScheme, Color.FromArgb(255, 0, 0)); // Red
                                accent2Prop.SetValue(colorScheme, Color.FromArgb(0, 0, 255)); // Blue
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    // If the Theme API is unavailable, continue without modifying theme colors
                    Console.WriteLine($"Theme modification skipped: {ex.Message}");
                }

                // Save the workbook with the updated (or original) theme colors
                string outputPath = "ModifiedThemeChart.xlsx";

                // Ensure the directory exists before saving
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                try
                {
                    workbook.Save(outputPath);
                    Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
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
}
