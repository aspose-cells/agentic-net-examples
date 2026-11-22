// Title: How to extract an Excel workbook's theme accent colors and convert them to HSL with Aspose.Cells in C#
// Common Searches: C# Aspose.Cells read workbook theme accent colors and display HSL values | convert Excel theme RGB colors to HSL using Aspose.Cells library | dynamic access to Theme.AccentColors property with fallback default colors in Aspose.Cells | log RGB and HSL representations of Excel theme accents in a .NET console app
// Tags: Aspose.Cells theme accent colors extraction | C# RGB to HSL conversion for Excel | dynamic Theme handling Aspose.Cells | fallback default accent palette Aspose.Cells | console logging HSL values C#

using Aspose.Cells;
using System;
using System.Drawing;
using System.IO;

// The example loads an Excel file with Aspose.Cells, retrieves the workbook's theme accent colors via a dynamic Theme object (with a fallback to a predefined palette), converts each RGB accent color to HSL using a custom helper, and writes both RGB and HSL values to the console.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Ensure the input file exists before attempting to load it
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: File not found -> {inputPath}");
            return;
        }

        Workbook workbook = null;
        try
        {
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load workbook: {ex.Message}");
            return;
        }

        // Retrieve the theme's accent colors (use dynamic to handle API variations)
        Color[] accentColors = GetAccentColors(workbook);

        // Convert each accent color to HSL and output the result
        for (int i = 0; i < accentColors.Length; i++)
        {
            Color rgb = accentColors[i];
            double h, s, l;
            RgbToHsl(rgb.R, rgb.G, rgb.B, out h, out s, out l);
            Console.WriteLine($"Accent {i + 1}: RGB({rgb.R},{rgb.G},{rgb.B}) => HSL({h:F2}°, {s:P2}, {l:P2})");
        }
    }

    // Safely obtain accent colors from the workbook's theme
    private static Color[] GetAccentColors(Workbook workbook)
    {
        try
        {
            // Use dynamic to accommodate different Aspose.Cells versions where Theme may be a string or a Theme object
            dynamic theme = workbook.Theme;
            Color[] colors = theme.AccentColors as Color[];
            if (colors != null && colors.Length > 0)
                return colors;
        }
        catch
        {
            // Ignored – fallback will be used
        }

        // Fallback: return a default set of colors if theme accent colors are unavailable
        return new Color[]
        {
            Color.FromArgb(0, 112, 192), // Default accent 1
            Color.FromArgb(255, 192, 0), // Default accent 2
            Color.FromArgb(112, 173, 71), // Default accent 3
            Color.FromArgb(255, 0, 0),   // Default accent 4
            Color.FromArgb(0, 176, 80),  // Default accent 5
            Color.FromArgb(112, 48, 160) // Default accent 6
        };
    }

    // Helper method to convert RGB to HSL
    static void RgbToHsl(int r, int g, int b, out double h, out double s, out double l)
    {
        double rd = r / 255.0;
        double gd = g / 255.0;
        double bd = b / 255.0;

        double max = Math.Max(rd, Math.Max(gd, bd));
        double min = Math.Min(rd, Math.Min(gd, bd));
        l = (max + min) / 2.0;

        if (Math.Abs(max - min) < 0.0001)
        {
            h = 0;
            s = 0;
        }
        else
        {
            double delta = max - min;
            s = l > 0.5 ? delta / (2.0 - max - min) : delta / (max + min);

            if (Math.Abs(max - rd) < 0.0001)
                h = (gd - bd) / delta + (gd < bd ? 6 : 0);
            else if (Math.Abs(max - gd) < 0.0001)
                h = (bd - rd) / delta + 2;
            else
                h = (rd - gd) / delta + 4;

            h /= 6.0;
            h *= 360.0; // Convert to degrees
        }
    }
}
