// Title: C# Example: Extract Workbook Theme Accent Colors with Aspose.Cells, Convert to HSL, and Log Results
// Description: This Aspose.Cells for .NET sample creates a new workbook, reads the six default theme accent colors (Accent1‑Accent6) using the GetThemeColor API, converts each RGB value to HSL with a lightweight routine, writes the hue, saturation and lightness to the console, and optionally saves the file. Ideal for developers who need to work with theme palettes, perform color transformations, or validate design consistency in Excel files.
// Keywords: Aspose.Cells theme colors C# | GetThemeColor Aspose.Cells | RGB to HSL conversion .NET | extract workbook accent colors | log HSL values console | default Excel theme palette | C# .NET Excel color handling | GitHub Aspose.Cells example
// Common Searches: how to read Excel theme accent colors with Aspose.Cells | convert workbook theme RGB to HSL in C# | Aspose.Cells GetThemeColor example | C# code to log HSL values of Excel theme colors | extract default theme palette Aspose.Cells
// Developer Intent: Read the workbook's theme accent colors, transform each from RGB to HSL, and output the HSL components.
// Use Cases: Build a custom color scheme by converting the default theme accents to HSL for further manipulation. | Verify that an Excel file follows branding guidelines by comparing logged HSL values against expected ranges. | Supply HSL color data to charting or graphics libraries that require hue, saturation, and lightness inputs.
// AI Prompts: Generate a C# method that returns a Dictionary<ThemeColorType, HslColor> containing all accent colors for a given Aspose.Cells workbook. | Provide code to export the HSL values of the six theme accent colors to a CSV file using Aspose.Cells. | Explain how to extend the sample to also convert background, text, and hyperlink theme colors to HSL.

using System;
using Aspose.Cells;
using System.Drawing;

// This Aspose.Cells for .NET sample creates a new workbook, reads the six default theme accent colors (Accent1‑Accent6) using the GetThemeColor API, converts each RGB value to HSL with a lightweight routine, writes the hue, saturation and lightness to the console, and optionally saves the file. Ideal for developers who need to work with theme palettes, perform color transformations, or validate design consistency in Excel files.
class ThemeAccentColorsHSL
{
    static void Main()
    {
        // Create a new workbook (default theme is applied)
        Workbook workbook = new Workbook();

        // Iterate through Accent1 to Accent6 theme colors
        for (int i = (int)ThemeColorType.Accent1; i <= (int)ThemeColorType.Accent6; i++)
        {
            ThemeColorType type = (ThemeColorType)i;
            Color rgb = workbook.GetThemeColor(type);          // Get the RGB color for the theme type
            HslColor hsl = RgbToHsl(rgb);                     // Convert to HSL
            Console.WriteLine($"{type}: H={hsl.H:F2}°, S={hsl.S:F2}%, L={hsl.L:F2}%");
        }

        // Save the workbook (optional, demonstrates lifecycle compliance)
        workbook.Save("ThemeAccentColors.xlsx");
    }

    // Simple struct to hold HSL components
    struct HslColor
    {
        public double H; // Hue in degrees (0‑360)
        public double S; // Saturation in percent (0‑100)
        public double L; // Lightness in percent (0‑100)

        public HslColor(double h, double s, double l)
        {
            H = h;
            S = s;
            L = l;
        }
    }

    // Convert a System.Drawing.Color (RGB) to HSL representation
    static HslColor RgbToHsl(Color color)
    {
        double r = color.R / 255.0;
        double g = color.G / 255.0;
        double b = color.B / 255.0;

        double max = Math.Max(r, Math.Max(g, b));
        double min = Math.Min(r, Math.Min(g, b));
        double h = 0.0, s, l = (max + min) / 2.0;

        if (Math.Abs(max - min) < 0.0001) // Achromatic case
        {
            s = 0.0;
        }
        else
        {
            double delta = max - min;
            s = l > 0.5 ? delta / (2.0 - max - min) : delta / (max + min);

            if (Math.Abs(max - r) < 0.0001)
                h = (g - b) / delta + (g < b ? 6.0 : 0.0);
            else if (Math.Abs(max - g) < 0.0001)
                h = (b - r) / delta + 2.0;
            else
                h = (r - g) / delta + 4.0;

            h /= 6.0;
        }

        // Convert to conventional HSL ranges
        return new HslColor(h * 360.0, s * 100.0, l * 100.0);
    }
}
