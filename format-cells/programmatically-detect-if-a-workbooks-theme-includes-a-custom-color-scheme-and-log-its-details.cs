// Title: Detect custom Excel theme and list its colors with Aspose.Cells for .NET
// Description: Shows how to read a workbook's Theme property, determine if it differs from the default "Office" theme, and enumerate all ThemeColorType values (except StyleColor) to log their ARGB components. Includes optional code to apply a custom theme and save the file.
// Keywords: Aspose.Cells custom theme detection | Excel theme colors .NET | GetThemeColor C# | ThemeColorType enumeration | Workbook.Theme property | log Excel theme colors
// Common Searches: how to check if an Excel file uses a custom theme with Aspose.Cells | retrieve theme colors from a workbook using C# | Aspose.Cells detect non‑Office theme | list ARGB values of Excel theme colors | C# code to read Excel theme scheme
// Developer Intent: Identify whether a workbook uses a non‑default theme and output each theme color's ARGB values.
// Use Cases: Verify that generated reports follow corporate branding by confirming the applied theme. | Audit a collection of spreadsheets for unauthorized custom themes. | Debug visual inconsistencies by logging exact theme color values.
// AI Prompts: Generate C# code with Aspose.Cells that checks if Workbook.Theme is not "Office" and prints all ThemeColorType colors with ARGB values. | Create a method that receives a Workbook and returns a dictionary of ThemeColorType to Color, excluding StyleColor, and a flag indicating a custom theme. | Provide an example that loads an existing .xlsx, detects a custom theme, and writes a text log of each theme color.

using System;
using System.Drawing;
using Aspose.Cells;

// Shows how to read a workbook's Theme property, determine if it differs from the default "Office" theme, and enumerate all ThemeColorType values (except StyleColor) to log their ARGB components. Includes optional code to apply a custom theme and save the file.
class DetectCustomTheme
{
    static void Main()
    {
        // Create a new workbook (you could also load an existing file)
        Workbook workbook = new Workbook();

        // ------------------------------------------------------------
        // OPTIONAL: Apply a custom theme to demonstrate detection.
        // In real scenarios you would load a workbook that may already
        // contain a custom theme.
        // ------------------------------------------------------------
        Color[] customColors = new Color[12];
        customColors[0] = Color.Red;          // Background1
        customColors[1] = Color.Green;        // Text1
        customColors[2] = Color.Blue;         // Background2
        customColors[3] = Color.Yellow;       // Text2
        // Remaining accent and hyperlink colors – using a neutral gray
        for (int i = 4; i < 12; i++)
        {
            customColors[i] = Color.FromArgb(255, 128, 128, 128);
        }
        workbook.CustomTheme("MyCustomTheme", customColors);

        // ------------------------------------------------------------
        // Detect whether the workbook uses a custom theme.
        // The default theme name in Excel is typically "Office".
        // If the name differs, we treat it as a custom theme.
        // ------------------------------------------------------------
        string themeName = workbook.Theme;
        bool isCustomTheme = !string.Equals(themeName, "Office", StringComparison.OrdinalIgnoreCase);

        Console.WriteLine($"Theme name: {themeName}");
        Console.WriteLine($"Is custom theme: {isCustomTheme}");

        // ------------------------------------------------------------
        // Log details of all theme colors.
        // ThemeColorType.StyleColor is internal and can be skipped.
        // ------------------------------------------------------------
        foreach (ThemeColorType type in Enum.GetValues(typeof(ThemeColorType)))
        {
            if (type == ThemeColorType.StyleColor) continue;

            Color color = workbook.GetThemeColor(type);
            Console.WriteLine($"{type}: A={color.A}, R={color.R}, G={color.G}, B={color.B}");
        }

        // Save the workbook (optional, demonstrates persistence of the theme)
        workbook.Save("DetectCustomTheme.xlsx");
    }
}
