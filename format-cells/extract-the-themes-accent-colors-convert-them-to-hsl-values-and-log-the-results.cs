// Title: Read Excel theme accent colors and output their HSL values with Aspose.Cells for .NET (C#)
// Common Searches: how to get theme accent colors from an .xlsx file using Aspose.Cells in C# | convert workbook theme RGB colors to HSL with Aspose.Cells .NET | sample code for extracting Excel theme colors and printing HSL in C# | Aspose.Cells GetThemeColor example converting to HSL format | retrieve and log six accent colors from Excel theme using C#
// Tags: Aspose.Cells GetThemeColor API | Excel theme accent color extraction C# | RGB to HSL conversion method C# | log theme colors as HSL Aspose.Cells | process workbook theme colors .NET

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// The example loads an .xlsx workbook, uses Workbook.GetThemeColor to fetch the six theme accent colors (Accent1‑Accent6), converts each System.Drawing.Color from RGB to HSL with a custom conversion routine, and writes both the original RGB and the calculated HSL values to the console.
class ThemeAccentExtractor
{
    static void Main()
    {
        const string filePath = "input.xlsx";

        // Verify that the input file exists
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"Error: File not found -> {filePath}");
            return;
        }

        Workbook workbook;
        try
        {
            // Load the workbook
            workbook = new Workbook(filePath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading workbook: {ex.Message}");
            return;
        }

        try
        {
            // Define the accent theme color types
            ThemeColorType[] accentTypes = new ThemeColorType[]
            {
                ThemeColorType.Accent1,
                ThemeColorType.Accent2,
                ThemeColorType.Accent3,
                ThemeColorType.Accent4,
                ThemeColorType.Accent5,
                ThemeColorType.Accent6
            };

            // Retrieve and process each accent color
            for (int i = 0; i < accentTypes.Length; i++)
            {
                Color rgb = workbook.GetThemeColor(accentTypes[i]);
                var hsl = RgbToHsl(rgb);
                Console.WriteLine(
                    $"Accent {i + 1}: RGB({rgb.R},{rgb.G},{rgb.B}) => HSL({hsl.H:F2}°, {hsl.S:P2}, {hsl.L:P2})");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error processing theme colors: {ex.Message}");
        }
    }

    // Convert a System.Drawing.Color (RGB) to HSL
    private static (double H, double S, double L) RgbToHsl(Color color)
    {
        double r = color.R / 255.0;
        double g = color.G / 255.0;
        double b = color.B / 255.0;

        double max = Math.Max(r, Math.Max(g, b));
        double min = Math.Min(r, Math.Min(g, b));
        double delta = max - min;

        // Lightness
        double l = (max + min) / 2.0;

        double h = 0.0;
        double s = 0.0;

        if (delta != 0)
        {
            // Saturation
            s = l < 0.5 ? delta / (max + min) : delta / (2.0 - max - min);

            // Hue
            if (max == r)
                h = ((g - b) / delta) % 6;
            else if (max == g)
                h = ((b - r) / delta) + 2;
            else // max == b
                h = ((r - g) / delta) + 4;

            h *= 60; // Convert to degrees
            if (h < 0)
                h += 360;
        }

        return (h, s, l);
    }
}
