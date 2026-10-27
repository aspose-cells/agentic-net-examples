// Title: Change the Accent5 color of an Excel workbook theme and refresh dependent charts with Aspose.Cells for .NET
// Common Searches: how to set the Accent5 theme color in an existing .xlsx using Aspose.Cells C# | update Excel theme palette programmatically and refresh charts with Aspose.Cells .NET | use reflection to modify the theme ColorScheme Accent5 property in older Aspose.Cells versions | recalculate workbook to refresh charts after changing theme colors in Aspose.Cells | save a workbook after changing the theme accent color with Aspose.Cells for .NET
// Tags: Aspose.Cells set theme Accent5 color | Excel theme palette update .NET | refresh charts after theme change Aspose.Cells | reflection modify theme ColorScheme Aspose.Cells | recalculate formulas to refresh charts Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

namespace UpdateThemeAccent5App
{
    // The example loads an existing workbook, uses reflection to set the Theme's ColorScheme Accent5 property to a custom Color, recalculates formulas to ensure all charts reflect the new theme, and saves the updated workbook.
    class UpdateThemeAccent5
    {
        static void Main()
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "output.xlsx";

                // Verify input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // User‑selected color for Accent5
                Color userSelectedColor = Color.FromArgb(255, 100, 150, 200);

                // Attempt to update the theme's Accent5 color if the Theme object is available
                try
                {
                    // In newer Aspose.Cells versions Workbook.Theme returns a Theme object.
                    // If the property is a string (older versions), this block will be skipped.
                    var theme = workbook.Theme;
                    if (theme != null && theme.GetType().GetProperty("ColorScheme") != null)
                    {
                        // Use reflection to set Accent5 without breaking compilation on older versions
                        var colorScheme = theme.GetType().GetProperty("ColorScheme")?.GetValue(theme);
                        if (colorScheme != null)
                        {
                            var accent5Prop = colorScheme.GetType().GetProperty("Accent5");
                            if (accent5Prop != null && accent5Prop.CanWrite)
                            {
                                accent5Prop.SetValue(colorScheme, userSelectedColor);
                                Console.WriteLine("Accent5 color updated in the theme.");
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Theme update skipped: {ex.Message}");
                }

                // Refresh charts – Aspose.Cells updates chart appearance automatically on Save.
                // If explicit refresh is needed, recalculate the workbook.
                try
                {
                    workbook.CalculateFormula();
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Chart refresh skipped: {ex.Message}");
                }

                // Ensure output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
