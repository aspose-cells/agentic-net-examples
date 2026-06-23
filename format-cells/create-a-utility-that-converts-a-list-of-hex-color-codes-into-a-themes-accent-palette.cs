using System;
using System.Collections.Generic;
using System.Drawing;
using Aspose.Cells;

public class ThemeUtility
{
    // Converts a list of hex color strings to a custom theme where the hex colors
    // populate the Accent1‑Accent6 slots of the theme palette.
    public static void CreateAccentTheme(List<string> hexColors, string themeName, string outputPath)
    {
        // Create a new workbook (lifecycle create)
        Workbook workbook = new Workbook();

        // Prepare a 12‑element array for the theme colors.
        // Index mapping:
        // 0‑3 : Background1, Text1, Background2, Text2 (use defaults)
        // 4‑9 : Accent1‑Accent6 (will be filled from hexColors)
        // 10‑11 : Hyperlink, FollowedHyperlink (use defaults)
        Color[] themeColors = new Color[12];

        // Fill the first four slots with the workbook's current theme colors as defaults.
        themeColors[0] = workbook.GetThemeColor(ThemeColorType.Background1);
        themeColors[1] = workbook.GetThemeColor(ThemeColorType.Text1);
        themeColors[2] = workbook.GetThemeColor(ThemeColorType.Background2);
        themeColors[3] = workbook.GetThemeColor(ThemeColorType.Text2);

        // Fill Accent slots (indices 4‑9) with provided hex colors.
        // If fewer than 6 colors are supplied, remaining accents keep the default theme colors.
        for (int i = 0; i < Math.Min(hexColors.Count, 6); i++)
        {
            // Convert hex string (e.g., "#FF5733" or "FF5733") to System.Drawing.Color.
            string hex = hexColors[i];
            if (!hex.StartsWith("#"))
                hex = "#" + hex;
            Color accentColor = ColorTranslator.FromHtml(hex);
            themeColors[4 + i] = accentColor;
        }

        // Fill remaining Accent slots with defaults if they were not set above.
        for (int i = hexColors.Count; i < 6; i++)
        {
            themeColors[4 + i] = workbook.GetThemeColor((ThemeColorType)(4 + i));
        }

        // Fill Hyperlink and FollowedHyperlink with defaults.
        themeColors[10] = workbook.GetThemeColor(ThemeColorType.Hyperlink);
        themeColors[11] = workbook.GetThemeColor(ThemeColorType.FollowedHyperlink);

        // Apply the custom theme (lifecycle operation).
        workbook.CustomTheme(themeName, themeColors);

        // Save the workbook (lifecycle save).
        workbook.Save(outputPath);
    }
}

// Example usage
class Program
{
    static void Main()
    {
        List<string> hexPalette = new List<string>
        {
            "FF5733", // Accent1
            "33FF57", // Accent2
            "3357FF", // Accent3
            "FF33A1", // Accent4
            "A1FF33", // Accent5
            "33A1FF"  // Accent6
        };

        ThemeUtility.CreateAccentTheme(hexPalette, "MyHexTheme", "HexThemeWorkbook.xlsx");
        Console.WriteLine("Custom theme created and workbook saved.");
    }
}