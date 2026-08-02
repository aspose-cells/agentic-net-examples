// Title: C# – Convert Hex Color Array to Excel Accent Theme with Aspose.Cells
// Description: A C# utility that builds a new Workbook, parses up to six hex color strings into the Accent1‑Accent6 slots of a 12‑color Excel theme, fills remaining slots with neutral defaults, applies the palette via workbook.CustomTheme, and saves the file as XLSX.
// Keywords: Aspose.Cells | C# custom Excel theme | hex color palette | Accent1 Accent6 | CustomTheme method | programmatic Excel styling | theme colors from hex | Excel workbook generation | brand colors in Excel | Aspose.Cells API
// Common Searches: How to create a custom Excel theme from hex colors using Aspose.Cells C# | Convert hex color codes to Excel accent palette programmatically | Aspose.Cells CustomTheme example C# | Set Excel theme colors from an array of hex strings | Apply brand color palette to Excel workbook with Aspose.Cells
// Developer Intent: Generate an Excel workbook that uses a custom theme whose Accent1‑Accent6 colors are defined by a supplied list of hex codes.
// Use Cases: Brand‑consistent reports: turn corporate hex palette into a reusable Excel theme for automated dashboards. | Template generation: create spreadsheet templates with predefined accent colors driven by user‑provided hex values. | Marketing analytics: standardize hyperlink colors while customizing accent shades for visual data stories.
// AI Prompts: Write a C# method that accepts up to six hex color strings, builds the 12‑element Color array required by Aspose.Cells, applies it as a custom theme, and saves the workbook. | Explain the mapping of each index in the Color[] passed to workbook.CustomTheme to the corresponding theme slots in an Excel file. | Add robust validation for hex strings and error handling to the theme‑creation utility using Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;

namespace ThemeUtility
{
    // A C# utility that builds a new Workbook, parses up to six hex color strings into the Accent1‑Accent6 slots of a 12‑color Excel theme, fills remaining slots with neutral defaults, applies the palette via workbook.CustomTheme, and saves the file as XLSX.
    public static class ThemeConverter
    {
        /// <param name="hexColors">Array of hex strings (e.g., "#FF5733"). Only the first six are used for Accent1‑Accent6.</param>
        /// <param name="themeName">Name of the custom theme to create.</param>
        /// <param name="outputPath">File path where the workbook will be saved.</param>
        public static void ConvertHexToAccentTheme(string[] hexColors, string themeName, string outputPath)
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Prepare a 12‑element array for the theme colors.
            // Indices 0‑3 and 10‑11 are non‑accent slots; we fill them with neutral colors.
            Color[] themeColors = new Color[12];
            themeColors[0] = Color.White;   // Background1
            themeColors[1] = Color.Black;   // Text1
            themeColors[2] = Color.White;   // Background2
            themeColors[3] = Color.Black;   // Text2

            // Fill Accent1‑Accent6 (indices 4‑9) with the supplied hex colors.
            int accentCount = Math.Min(6, hexColors.Length);
            for (int i = 0; i < accentCount; i++)
            {
                // Parse hex string to System.Drawing.Color.
                // ColorTranslator handles strings with or without leading '#'.
                Color parsed = ColorTranslator.FromHtml(hexColors[i]);
                themeColors[4 + i] = parsed;
            }

            // If fewer than six colors were supplied, fill remaining accents with a default gray.
            for (int i = accentCount; i < 6; i++)
            {
                themeColors[4 + i] = Color.Gray;
            }

            // Hyperlink and FollowedHyperlink slots (indices 10‑11) get default colors.
            themeColors[10] = Color.Blue;          // Hyperlink
            themeColors[11] = Color.Purple;        // Followed Hyperlink

            // Apply the custom theme (rule: CustomTheme)
            workbook.CustomTheme(themeName, themeColors);

            // Save the workbook (lifecycle rule: save)
            workbook.Save(outputPath, SaveFormat.Xlsx);
        }
    }

    // Example usage
    class Program
    {
        static void Main()
        {
            string[] hexPalette = new string[]
            {
                "#FF5733", // Accent1
                "#33FF57", // Accent2
                "#3357FF", // Accent3
                "#FF33A1", // Accent4
                "#A133FF", // Accent5
                "#33FFF5"  // Accent6
            };

            ThemeConverter.ConvertHexToAccentTheme(
                hexPalette,
                "MyHexTheme",
                "HexThemeWorkbook.xlsx"
            );

            Console.WriteLine("Custom theme created and workbook saved.");
        }
    }
}
