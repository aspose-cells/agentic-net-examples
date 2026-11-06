// Title: Create an Excel workbook with a custom theme accent palette from hex color codes using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells set custom theme accent colors from a list of hex values | How to generate an Excel theme palette using hex color codes with Aspose.Cells | Example of Workbook.SetThemeColor for Accent1‑Accent6 in .NET | Create Excel workbook with custom theme colors programmatically using Aspose.Cells
// Tags: Aspose.Cells theme accent color assignment | C# hex to Excel theme conversion | Workbook.SetThemeColor usage | custom Excel theme palette generation | apply hex palette to workbook theme

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing; // For ThemeColorType enum

namespace ThemePaletteUtility
{
    // This utility takes up to six hex color strings, creates a new Workbook, maps each color to the corresponding ThemeColorType (Accent1‑Accent6) via Workbook.SetThemeColor, ensures the output directory exists, and saves the workbook to the specified path while handling errors.
    public static class ThemeConverter
    {
        /// <param name="hexColors">List of hex color strings (e.g., "#FF5733"). Expected up to 6 colors for Accent1‑Accent6.</param>
        /// <param name="outputFilePath">Full path where the workbook will be saved.</param>
        public static void ConvertHexToThemePalette(List<string> hexColors, string outputFilePath)
        {
            try
            {
                // Ensure the output directory exists
                string? directory = Path.GetDirectoryName(outputFilePath);
                if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                // Create a new empty workbook
                Workbook workbook = new Workbook();

                // Map each accent index to its corresponding ThemeColorType
                ThemeColorType[] accentTypes = new ThemeColorType[]
                {
                    ThemeColorType.Accent1,
                    ThemeColorType.Accent2,
                    ThemeColorType.Accent3,
                    ThemeColorType.Accent4,
                    ThemeColorType.Accent5,
                    ThemeColorType.Accent6
                };

                // Apply provided hex colors to the theme accents
                for (int i = 0; i < Math.Min(hexColors.Count, accentTypes.Length); i++)
                {
                    string hex = hexColors[i];

                    // Ensure the hex string starts with '#'
                    if (!hex.StartsWith("#"))
                    {
                        hex = "#" + hex;
                    }

                    // Convert hex string to System.Drawing.Color
                    Color color = ColorTranslator.FromHtml(hex);

                    // Set the corresponding accent color in the workbook theme
                    workbook.SetThemeColor(accentTypes[i], color);
                }

                // Save the workbook
                workbook.Save(outputFilePath);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error converting hex colors to theme palette: {ex.Message}");
                throw;
            }
        }
    }

    // Example usage
    class Program
    {
        static void Main()
        {
            try
            {
                List<string> hexPalette = new List<string>
                {
                    "#1F77B4", // Accent1
                    "#FF7F0E", // Accent2
                    "#2CA02C", // Accent3
                    "#D62728", // Accent4
                    "#9467BD", // Accent5
                    "#8C564B"  // Accent6
                };

                string outputPath = "ThemePaletteWorkbook.xlsx";

                ThemeConverter.ConvertHexToThemePalette(hexPalette, outputPath);

                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }
}
