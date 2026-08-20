// Title: C# utility to set Excel theme accent palette from hex colors using Aspose.Cells
// Description: A reusable C# method that builds a 12‑color Excel theme array, preserves default background, text and hyperlink colors, converts up to six hex strings ("#RRGGBB" or "RRGGBB") to System.Drawing.Color, assigns them to Accent1‑Accent6, applies the custom theme with Workbook.CustomTheme, and saves the workbook as XLSX.
// Keywords: Aspose.Cells | C# | .NET | Excel custom theme | accent palette | hex color to Excel theme | ColorTranslator | Workbook.CustomTheme | programmatic Excel styling | brand colors in Excel | theme colors API
// Common Searches: Aspose.Cells set custom accent colors C# | convert hex codes to Excel theme palette .NET | apply brand color palette to Excel workbook using Aspose | how to change Accent1‑Accent6 programmatically in Excel | C# example for custom Excel theme with Aspose.Cells
// Developer Intent: Create or modify an Excel workbook so that its Accent1‑Accent6 colors come from a supplied list of hex codes.
// Use Cases: Generate corporate‑branded reports by mapping a company’s hex palette to Excel theme accents. | Provide a UI where users pick colors, then output a themed spreadsheet that reflects those selections. | Automate production of department‑specific templates, each with a distinct accent color set.
// AI Prompts: Write a C# function that receives a List<string> of hex colors and applies them to Accent1‑Accent6 in a custom Excel theme with Aspose.Cells. | Explain how to safely parse hex strings to System.Drawing.Color and fallback to default theme colors when parsing fails. | Show how to reuse the custom theme utility to create multiple workbooks with different theme names and output file paths.

using System;
using System.Collections.Generic;
using System.Drawing;
using Aspose.Cells;

namespace ThemeUtilityDemo
{
    // A reusable C# method that builds a 12‑color Excel theme array, preserves default background, text and hyperlink colors, converts up to six hex strings ("#RRGGBB" or "RRGGBB") to System.Drawing.Color, assigns them to Accent1‑Accent6, applies the custom theme with Workbook.CustomTheme, and saves the workbook as XLSX.
    public static class ThemeUtility
    {
        /// <param name="hexColors">Hex strings (e.g., "#FF1122" or "FF1122") for Accent1‑Accent6. Up to 6 colors are used.</param>
        /// <param name="themeName">Name of the custom theme.</param>
        /// <param name="outputPath">File path where the workbook will be saved.</param>
        public static void ApplyAccentPalette(List<string> hexColors, string themeName, string outputPath)
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Prepare the 12 theme colors array.
            // Index mapping (see Workbook.CustomTheme documentation):
            // 0‑3 : Background1, Text1, Background2, Text2 (use existing theme defaults)
            // 4‑9 : Accent1‑Accent6 (filled from hexColors)
            // 10‑11 : Hyperlink, FollowedHyperlink (use existing theme defaults)
            Color[] themeColors = new Color[12];

            // Fill the first four entries with the current theme's defaults.
            themeColors[0] = workbook.GetThemeColor(ThemeColorType.Background1);
            themeColors[1] = workbook.GetThemeColor(ThemeColorType.Text1);
            themeColors[2] = workbook.GetThemeColor(ThemeColorType.Background2);
            themeColors[3] = workbook.GetThemeColor(ThemeColorType.Text2);

            // Populate Accent1‑Accent6 from the supplied hex list.
            // If fewer than 6 colors are supplied, remaining accents keep the default.
            for (int i = 0; i < 6; i++)
            {
                int themeIndex = 4 + i; // Accent1 starts at index 4
                if (i < hexColors.Count)
                {
                    // Convert hex string to System.Drawing.Color.
                    // ColorTranslator handles both "#RRGGBB" and "RRGGBB".
                    try
                    {
                        themeColors[themeIndex] = ColorTranslator.FromHtml(hexColors[i]);
                    }
                    catch
                    {
                        // If conversion fails, fall back to the existing theme color.
                        themeColors[themeIndex] = workbook.GetThemeColor((ThemeColorType)themeIndex);
                    }
                }
                else
                {
                    // Use existing theme accent if no custom value is provided.
                    themeColors[themeIndex] = workbook.GetThemeColor((ThemeColorType)themeIndex);
                }
            }

            // Fill Hyperlink and FollowedHyperlink with current theme defaults.
            themeColors[10] = workbook.GetThemeColor(ThemeColorType.Hyperlink);
            themeColors[11] = workbook.GetThemeColor(ThemeColorType.FollowedHyperlink);

            // Apply the custom theme (lifecycle rule: modify)
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
            List<string> hexPalette = new List<string>
            {
                "#4F81BD", // Accent1
                "#C0504D", // Accent2
                "#9BBB59", // Accent3
                "#8064A2", // Accent4
                "#4BACC6", // Accent5
                "#F79646"  // Accent6
            };

            ThemeUtility.ApplyAccentPalette(hexPalette, "MyCustomTheme", "CustomThemeDemo.xlsx");
            Console.WriteLine("Custom theme applied and workbook saved.");
        }
    }
}
