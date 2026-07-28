// Title: Extract Excel Theme Accent Colors and Convert to HSL with Aspose.Cells for .NET (C#)
// Description: This C# example creates a new Workbook, reads the six default theme accent colors via GetThemeColor, converts each RGB value to HSL using a custom method, and writes the RGB and HSL values to the console.
// Keywords: Aspose.Cells | C# Excel theme colors | GetThemeColor | ThemeColorType Accent1 | RGB to HSL conversion | extract Excel accent colors | default workbook theme | color conversion .NET | Excel theme palette | HSL values C#
// Common Searches: Aspose.Cells get theme accent colors C# | Convert Excel theme colors to HSL | How to read Excel theme palette with Aspose | C# code to convert RGB to HSL | GetThemeColor example Aspose.Cells | Extract default theme colors from workbook | Log theme accent colors in .NET
// Developer Intent: Read the workbook’s theme accent palette, transform each RGB entry to HSL, and output the results.
// Use Cases: Build a custom color picker that works with HSL sliders | Create charts with gradient fills that require HSL inputs | Automate brand‑compliance checks by comparing extracted HSL values to corporate standards | Generate CSS or design tokens from Excel theme colors | Export theme colors to JSON for front‑end styling
// AI Prompts: Generate a C# function that returns all theme accent colors from an Aspose.Cells Workbook as a List<Color>. | Provide a minimal RGB‑to‑HSL conversion routine in C# without external libraries. | Show how to write the extracted RGB and HSL values to a CSV file instead of the console. | Explain how to map the HSL values to CSS hue‑rotate filters using Aspose.Cells data. | Create a PowerShell script that calls the compiled .NET assembly to list theme accent colors.

using System;
using System.Drawing;
using Aspose.Cells;

namespace ThemeAccentExtractor
{
    // This C# example creates a new Workbook, reads the six default theme accent colors via GetThemeColor, converts each RGB value to HSL using a custom method, and writes the RGB and HSL values to the console.
    class Program
    {
        // Convert an RGB color to HSL.
        // Returns hue (0-360), saturation (0-1), lightness (0-1).
        static (double H, double S, double L) RgbToHsl(Color color)
        {
            double r = color.R / 255.0;
            double g = color.G / 255.0;
            double b = color.B / 255.0;

            double max = Math.Max(r, Math.Max(g, b));
            double min = Math.Min(r, Math.Min(g, b));
            double delta = max - min;

            // Lightness
            double l = (max + min) / 2.0;

            // Saturation
            double s = 0;
            if (delta != 0)
            {
                s = l < 0.5
                    ? delta / (max + min)
                    : delta / (2.0 - max - min);
            }

            // Hue
            double h = 0;
            if (delta != 0)
            {
                if (max == r)
                {
                    h = ((g - b) / delta) % 6.0;
                }
                else if (max == g)
                {
                    h = ((b - r) / delta) + 2.0;
                }
                else // max == b
                {
                    h = ((r - g) / delta) + 4.0;
                }
                h *= 60.0;
                if (h < 0) h += 360.0;
            }

            return (h, s, l);
        }

        static void Main()
        {
            // Create a new workbook (default theme is applied)
            Workbook workbook = new Workbook();

            // Iterate over the six accent colors defined in the theme
            for (int i = 4; i <= 9; i++) // ThemeColorType.Accent1 = 4, Accent6 = 9
            {
                ThemeColorType accentType = (ThemeColorType)i;
                Color rgbColor = workbook.GetThemeColor(accentType);
                var (h, s, l) = RgbToHsl(rgbColor);

                Console.WriteLine($"Accent{i - 3}: RGB = {rgbColor} | HSL = ({h:F2}°, {s:P2}, {l:P2})");
            }
        }
    }
}
