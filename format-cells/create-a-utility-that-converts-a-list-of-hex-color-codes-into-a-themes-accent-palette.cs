// Title: Create an Excel workbook with a custom theme accent palette from hex color codes using Aspose.Cells for .NET
// AI Prompts: Write a C# method that accepts a file path and a List<string> of hex colors, converts each hex value to System.Drawing.Color, and assigns it to ThemeColorType.Accent1‑Accent6 using Workbook.SetThemeColor. | Add logic to verify the output directory exists, create it if missing, and save the workbook to the specified location with Aspose.Cells. | Show how to call the utility with six sample hex values (e.g., #FF5733) and save the themed workbook to the desktop.
// Common Searches: Aspose.Cells how to programmatically set theme accent colors from hex values in C# | C# convert hex color string to Excel theme palette using Aspose.Cells | Set custom theme colors for a new workbook with Aspose.Cells .NET example | Create workbook with user-defined accent palette using Aspose.Cells and System.Drawing.Color | Save themed Excel file to a specific folder after creating output directory in C#
// Tags: Aspose.Cells set custom theme accent palette | Workbook.SetThemeColor with hex colors | C# convert hex to System.Drawing.Color for Excel theme | Create themed Excel workbook .NET | Ensure output directory exists before saving workbook

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // The utility creates a new Workbook, ensures the target folder exists, converts each supplied hex string to System.Drawing.Color, maps the colors to ThemeColorType.Accent1‑Accent6 via Workbook.SetThemeColor, and saves the workbook to the given path.
    public static class ThemePaletteUtility
    {
        /// <param name="outputPath">Full path where the workbook will be saved.</param>
        /// <param name="hexColors">List of hex color strings (e.g., "#FF5733"). Up to six colors are used.</param>
        public static void CreateWorkbookWithAccentPalette(string outputPath, List<string> hexColors)
        {
            try
            {
                // Ensure the output directory exists
                string directory = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                // Create a new workbook
                Workbook workbook = new Workbook();

                // Map each hex color to the corresponding accent slot (Accent1‑Accent6)
                ThemeColorType[] accentTypes = new ThemeColorType[]
                {
                    ThemeColorType.Accent1,
                    ThemeColorType.Accent2,
                    ThemeColorType.Accent3,
                    ThemeColorType.Accent4,
                    ThemeColorType.Accent5,
                    ThemeColorType.Accent6
                };

                for (int i = 0; i < hexColors.Count && i < accentTypes.Length; i++)
                {
                    // Convert hex string to System.Drawing.Color
                    Color color = ColorTranslator.FromHtml(hexColors[i]);

                    // Apply the color to the workbook's theme accent
                    workbook.SetThemeColor(accentTypes[i], color);
                }

                // Save the workbook
                workbook.Save(outputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error creating workbook: {ex.Message}");
                throw;
            }
        }
    }

    // Simple entry point for demonstration
    public class Program
    {
        public static void Main()
        {
            try
            {
                var colors = new List<string>
                {
                    "#FF5733", "#33FF57", "#3357FF",
                    "#F1C40F", "#9B59B6", "#1ABC9C"
                };

                string outputPath = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                    "ThemedWorkbook.xlsx");

                ThemePaletteUtility.CreateWorkbookWithAccentPalette(outputPath, colors);
                Console.WriteLine($"Workbook saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }
}
