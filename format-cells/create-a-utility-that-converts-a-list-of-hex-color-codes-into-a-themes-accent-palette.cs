// Title: Apply a List of Hex Color Codes to an Excel Workbook’s Theme Accent Palette with Aspose.Cells for .NET (C#)
// Common Searches: c# aspose.cells set workbook theme accent colors from hex strings | how to populate Excel theme accent palette using a hex color array in Aspose.Cells | convert hex color list to theme accents in a .NET workbook with Aspose.Cells | dynamic assignment of theme.Accent1‑Accent6 properties in Aspose.Cells C# | save workbook after customizing theme colors using hex codes Aspose.Cells
// Tags: Aspose.Cells set theme accent palette C# | hex color to System.Drawing.Color conversion | custom Excel theme colors using Aspose.Cells | dynamic theme property access Aspose.Cells | populate workbook theme accents from list

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing; // Retained for potential future use

namespace ThemeUtility
{
    // The example shows how to take up to six hex color strings, convert each to a System.Drawing.Color, and assign them to the workbook's Theme Accent1‑Accent6 properties via dynamic access. It validates inputs, creates a new Workbook, applies the custom palette, ensures the output folder exists, saves the workbook as an .xlsx file, and logs the result.
    public static class ThemeHelper
    {
        /// <param name="hexColors">List of hex color strings (e.g., "#FF5733").</param>
        /// <param name="workbook">The Aspose.Cells workbook whose theme will be modified.</param>
        public static void ApplyHexColorsToThemePalette(List<string> hexColors, Workbook workbook)
        {
            if (hexColors == null) throw new ArgumentNullException(nameof(hexColors));
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));

            // Use dynamic to avoid compile‑time dependency on the Theme type (may vary by version).
            dynamic theme = workbook.Theme;

            int count = Math.Min(6, hexColors.Count);
            for (int i = 0; i < count; i++)
            {
                Color color = ColorTranslator.FromHtml(hexColors[i]);

                // Map each color to the corresponding accent property.
                switch (i)
                {
                    case 0: theme.Accent1 = color; break;
                    case 1: theme.Accent2 = color; break;
                    case 2: theme.Accent3 = color; break;
                    case 3: theme.Accent4 = color; break;
                    case 4: theme.Accent5 = color; break;
                    case 5: theme.Accent6 = color; break;
                }
            }
        }

        // Example usage.
        public static void Main()
        {
            try
            {
                // Sample list of hex color codes.
                List<string> hexColors = new List<string>
                {
                    "#FF5733", // Accent1
                    "#33FF57", // Accent2
                    "#3357FF", // Accent3
                    "#F1C40F", // Accent4
                    "#9B59B6", // Accent5
                    "#1ABC9C"  // Accent6
                };

                // Create a new workbook.
                Workbook workbook = new Workbook();

                // Apply the hex colors to the workbook's theme.
                ApplyHexColorsToThemePalette(hexColors, workbook);

                // Define output path.
                string outputPath = "ThemedWorkbook.xlsx";

                // Ensure the directory exists.
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook to verify the theme changes.
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
