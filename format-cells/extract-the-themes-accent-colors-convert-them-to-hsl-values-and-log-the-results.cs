using System;
using System.Drawing;
using Aspose.Cells;

namespace ThemeAccentColorsHSL
{
    class Program
    {
        // Convert a System.Drawing.Color to HSL components.
        // Returns a tuple: (Hue in degrees [0-360), Saturation [0-1], Lightness [0-1])
        static (double Hue, double Saturation, double Lightness) ColorToHsl(Color color)
        {
            // Normalize RGB values to the range 0-1.
            double r = color.R / 255.0;
            double g = color.G / 255.0;
            double b = color.B / 255.0;

            double max = Math.Max(r, Math.Max(g, b));
            double min = Math.Min(r, Math.Min(g, b));
            double delta = max - min;

            // Lightness calculation.
            double l = (max + min) / 2.0;

            double h = 0.0;
            double s = 0.0;

            if (delta != 0)
            {
                // Saturation calculation.
                s = l < 0.5 ? delta / (max + min) : delta / (2.0 - max - min);

                // Hue calculation.
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

                h *= 60.0; // Convert to degrees.
                if (h < 0)
                    h += 360.0;
            }

            return (h, s, l);
        }

        static void Main(string[] args)
        {
            // Create a new workbook (default theme is applied).
            Workbook workbook = new Workbook();

            // Array of accent theme color types.
            ThemeColorType[] accentTypes = new ThemeColorType[]
            {
                ThemeColorType.Accent1,
                ThemeColorType.Accent2,
                ThemeColorType.Accent3,
                ThemeColorType.Accent4,
                ThemeColorType.Accent5,
                ThemeColorType.Accent6
            };

            // Iterate through each accent color, retrieve the ARGB color,
            // convert it to HSL, and log the result.
            foreach (ThemeColorType accent in accentTypes)
            {
                Color rgbColor = workbook.GetThemeColor(accent);
                var (hue, saturation, lightness) = ColorToHsl(rgbColor);

                Console.WriteLine($"Accent {accent} - ARGB: {rgbColor.ToArgb():X8}");
                Console.WriteLine($"   HSL => Hue: {hue:F2}°, Saturation: {saturation:P1}, Lightness: {lightness:P1}");
            }

            // No need to save the workbook for this extraction task.
        }
    }
}