// Title: Merge Excel theme palettes with Aspose.Cells – primary workbook colors win (C#)
// Description: A C# utility that combines the theme palettes of two workbooks using Aspose.Cells. It copies the secondary workbook's theme, then restores the first 12 theme colors from the primary workbook so that primary colors retain precedence.
// Keywords: Aspose.Cells theme merge | C# Excel theme palette | CopyTheme method | SetThemeColor example | ThemeColorType enumeration | preserve primary colors | merge workbooks programmatically | Excel theme colors
// Common Searches: how to merge theme palettes in Aspose.Cells | preserve original theme colors when copying Excel theme | combine two workbook themes C# Aspose | CopyTheme without overwriting primary colors | set specific theme colors after copying workbook
// Developer Intent: Combine a secondary workbook's theme palette into a primary workbook while keeping the primary workbook's theme colors unchanged.
// Use Cases: Apply corporate brand colors and import additional accents from a department template without losing the brand palette. | Create a consolidated report that fills missing theme colors from a reference file while retaining any custom colors already defined. | Generate a master workbook that inherits all theme colors from a source workbook but preserves pre‑set primary accent colors.
// AI Prompts: Generate a C# method that merges two Aspose.Cells workbooks' theme palettes, giving priority to the primary workbook's colors. | Explain how to use CopyTheme and SetThemeColor together to merge Excel theme palettes without overwriting existing colors. | Write unit tests for ThemeMerger.MergeThemes covering overlapping and non‑overlapping ThemeColorType values.

using System;
using System.Collections.Generic;
using System.Drawing;
using Aspose.Cells;

// A C# utility that combines the theme palettes of two workbooks using Aspose.Cells. It copies the secondary workbook's theme, then restores the first 12 theme colors from the primary workbook so that primary colors retain precedence.
public static class ThemeMerger
{
    // Merges the theme palette of a secondary workbook into a primary workbook.
    // Colors defined in the primary workbook take precedence.
    public static void MergeThemes(Workbook primary, Workbook secondary)
    {
        // Store the original theme colors of the primary workbook.
        var originalColors = new Dictionary<ThemeColorType, Color>();
        foreach (ThemeColorType type in Enum.GetValues(typeof(ThemeColorType)))
        {
            // Only the first 12 types are actual theme colors.
            if ((int)type > 11) break;
            originalColors[type] = primary.GetThemeColor(type);
        }

        // Copy the theme from the secondary workbook to the primary workbook.
        // This brings in all theme colors from the secondary workbook.
        primary.CopyTheme(secondary);

        // Restore the primary workbook's original theme colors,
        // ensuring they have precedence over the copied ones.
        foreach (var kvp in originalColors)
        {
            primary.SetThemeColor(kvp.Key, kvp.Value);
        }
    }
}

// Example usage
public class Program
{
    public static void Main()
    {
        // Create primary workbook and customize some of its theme colors.
        Workbook primary = new Workbook();
        primary.SetThemeColor(ThemeColorType.Accent1, Color.FromArgb(255, 200, 0, 0)); // Dark red
        primary.SetThemeColor(ThemeColorType.Hyperlink, Color.Blue);

        // Create secondary workbook and customize a different set of theme colors.
        Workbook secondary = new Workbook();
        secondary.SetThemeColor(ThemeColorType.Accent1, Color.Green); // This will be overridden by primary
        secondary.SetThemeColor(ThemeColorType.Accent2, Color.Orange);
        secondary.SetThemeColor(ThemeColorType.Hyperlink, Color.Purple); // This will be overridden by primary

        // Merge the themes: primary colors win, secondary fills the rest.
        ThemeMerger.MergeThemes(primary, secondary);

        // Save the result to verify the merged theme.
        primary.Save("MergedThemeWorkbook.xlsx", SaveFormat.Xlsx);
    }
}
