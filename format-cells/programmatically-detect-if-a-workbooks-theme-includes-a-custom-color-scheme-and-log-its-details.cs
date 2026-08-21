// Title: Detect Custom Excel Theme and List Its Colors with Aspose.Cells for .NET
// Description: Load an Excel workbook, verify if its theme is not the default "Office" theme, and when a custom theme is present enumerate the 12 ThemeColorType slots via GetThemeColor, outputting each color's ARGB components. Includes optional workbook save.
// Keywords: Aspose.Cells | custom theme detection | enumerate theme colors | GetThemeColor | Excel workbook theme | C# | .NET | ThemeColorType | ARGB values
// Common Searches: Aspose.Cells check if workbook uses custom theme | list all theme colors from Excel file C# | retrieve ARGB values of Excel theme colors | detect non‑default theme in .NET Excel workbook | how to get theme color palette with Aspose.Cells
// Developer Intent: Identify whether a workbook uses a non‑default theme and extract the full set of theme colors.
// Use Cases: Validate corporate branding by confirming the workbook’s theme matches a predefined palette. | Create a design audit report that lists every theme color and its ARGB values for migration or compliance. | Drive conditional‑formatting logic that adapts based on the detected custom theme colors.
// AI Prompts: Generate a C# method that returns a Dictionary<ThemeColorType, Color> for any workbook, handling both default and custom themes. | Write a reusable Aspose.Cells utility class that detects a custom theme and logs each theme color with ARGB values, including file‑not‑found handling. | Provide sample code to compare extracted custom theme colors against a corporate color standard and flag any mismatches.

using System;
using System.Drawing;
using Aspose.Cells;

// Load an Excel workbook, verify if its theme is not the default "Office" theme, and when a custom theme is present enumerate the 12 ThemeColorType slots via GetThemeColor, outputting each color's ARGB components. Includes optional workbook save.
class DetectCustomTheme
{
    static void Main()
    {
        // Load an existing workbook (replace the path with your file)
        Workbook workbook = new Workbook("input.xlsx"); // load rule

        // Retrieve the theme name
        string themeName = workbook.Theme;
        Console.WriteLine($"Theme name: {themeName}");

        // Determine if the theme is custom.
        // The default theme name is usually "Office". Any other name indicates a custom theme.
        bool isCustom = !string.Equals(themeName, "Office", StringComparison.OrdinalIgnoreCase);
        Console.WriteLine($"Is custom theme: {isCustom}");

        // If a custom theme is present, enumerate its 12 theme colors.
        if (isCustom)
        {
            // ThemeColorType enum values 0‑11 correspond to the 12 theme slots.
            for (int i = 0; i <= 11; i++)
            {
                ThemeColorType type = (ThemeColorType)i;
                Color color = workbook.GetThemeColor(type);
                Console.WriteLine($"{type}: A={color.A}, R={color.R}, G={color.G}, B={color.B}");
            }
        }

        // Save the workbook (optional, demonstrates the save rule)
        workbook.Save("output.xlsx");
    }
}
