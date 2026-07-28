// Title: Extract Excel theme accent colors to a CSS file with Aspose.Cells for .NET
// Description: This example shows how to load or create a Workbook, read the six theme accent colors using GetThemeColor, convert them to HEX values, generate CSS custom properties (--accent1‑--accent6) with optional sample classes, and save the stylesheet as a .css file.
// Keywords: Aspose.Cells | C# | Excel theme colors | GetThemeColor | CSS variables | custom properties | generate stylesheet | web styling | color scheme export | theme accent extraction
// Common Searches: Aspose.Cells export theme colors to CSS | C# read Excel accent colors and create stylesheet | Generate CSS variables from Excel workbook theme | How to convert Excel theme colors to HEX in .NET | Create web stylesheet from Excel theme using Aspose
// Developer Intent: Read the workbook’s six accent colors and write them to a CSS file as variables.
// Use Cases: Synchronize website branding with colors defined in an Excel template. | Automatically refresh a shared CSS file when the Excel theme is updated. | Produce consistent color palettes for email or UI components directly from a spreadsheet.
// AI Prompts: Provide C# code that uses Aspose.Cells to read all six theme accent colors from a workbook and output a CSS file with custom properties and example classes. | Show how to modify the workbook’s theme accent colors before exporting them to a CSS stylesheet using Aspose.Cells. | Explain how to handle missing or default theme colors when generating a CSS file with Aspose.Cells.

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

namespace ThemeColorsToCss
{
    // This example shows how to load or create a Workbook, read the six theme accent colors using GetThemeColor, convert them to HEX values, generate CSS custom properties (--accent1‑--accent6) with optional sample classes, and save the stylesheet as a .css file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (or load an existing one that has a custom theme)
            Workbook workbook = new Workbook();

            // Example: modify some theme accent colors (optional)
            // workbook.SetThemeColor(ThemeColorType.Accent1, Color.FromArgb(255, 100, 150));
            // workbook.SetThemeColor(ThemeColorType.Accent2, Color.FromArgb(200, 50, 75));
            // ...

            // Retrieve the six accent colors from the workbook theme
            Color[] accentColors = new Color[6];
            accentColors[0] = workbook.GetThemeColor(ThemeColorType.Accent1);
            accentColors[1] = workbook.GetThemeColor(ThemeColorType.Accent2);
            accentColors[2] = workbook.GetThemeColor(ThemeColorType.Accent3);
            accentColors[3] = workbook.GetThemeColor(ThemeColorType.Accent4);
            accentColors[4] = workbook.GetThemeColor(ThemeColorType.Accent5);
            accentColors[5] = workbook.GetThemeColor(ThemeColorType.Accent6);

            // Build CSS content using CSS custom properties (variables)
            // This allows web developers to reference the theme colors easily.
            string css = ":root {\n";
            for (int i = 0; i < accentColors.Length; i++)
            {
                Color c = accentColors[i];
                // Convert the Color to a hex string like #RRGGBB
                string hex = $"#{c.R:X2}{c.G:X2}{c.B:X2}";
                css += $"    --accent{i + 1}: {hex};\n";
            }
            css += "}\n\n";

            // Optionally, provide sample CSS classes that use the variables
            css += ".accent1 { color: var(--accent1); }\n";
            css += ".accent2 { color: var(--accent2); }\n";
            css += ".accent3 { color: var(--accent3); }\n";
            css += ".accent4 { color: var(--accent4); }\n";
            css += ".accent5 { color: var(--accent5); }\n";
            css += ".accent6 { color: var(--accent6); }\n";

            // Define the output CSS file path
            string cssFilePath = "theme-colors.css";

            // Write the CSS content to the file
            File.WriteAllText(cssFilePath, css);

            Console.WriteLine($"CSS file with theme accent colors generated at: {Path.GetFullPath(cssFilePath)}");
        }
    }
}
