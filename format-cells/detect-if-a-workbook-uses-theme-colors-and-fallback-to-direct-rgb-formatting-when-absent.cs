using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsThemeFallback
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all used cells in the worksheet
                foreach (Cell cell in sheet.Cells)
                {
                    // Get the current style of the cell
                    Style style = cell.GetStyle();

                    // ----- Font Theme Color -----
                    if (style.Font.ThemeColor != null)
                    {
                        // Resolve the actual RGB color from the theme
                        Color baseColor = workbook.GetThemeColor(style.Font.ThemeColor.ColorType);
                        Color rgbColor = ApplyTint(baseColor, style.Font.ThemeColor.Tint);

                        // Replace theme color with direct RGB color
                        style.Font.Color = rgbColor;
                        style.Font.ThemeColor = null; // clear theme reference
                    }

                    // ----- Foreground Theme Color -----
                    if (style.ForegroundThemeColor != null)
                    {
                        Color baseColor = workbook.GetThemeColor(style.ForegroundThemeColor.ColorType);
                        Color rgbColor = ApplyTint(baseColor, style.ForegroundThemeColor.Tint);

                        style.ForegroundColor = rgbColor;
                        style.ForegroundThemeColor = null;
                    }

                    // ----- Background Theme Color -----
                    if (style.BackgroundThemeColor != null)
                    {
                        Color baseColor = workbook.GetThemeColor(style.BackgroundThemeColor.ColorType);
                        Color rgbColor = ApplyTint(baseColor, style.BackgroundThemeColor.Tint);

                        style.BackgroundColor = rgbColor;
                        style.BackgroundThemeColor = null;
                    }

                    // Apply the modified style back to the cell
                    cell.SetStyle(style);
                }
            }

            // Save the workbook with direct RGB formatting
            workbook.Save("output_fallback.xlsx");
        }

        // Applies Excel tint algorithm to a base color.
        // Positive tint lightens the color, negative tint darkens it.
        private static Color ApplyTint(Color baseColor, double tint)
        {
            // Clamp tint between -1 and 1
            tint = Math.Max(-1.0, Math.Min(1.0, tint));

            // Helper to apply tint to a single channel
            int ApplyChannel(int channel)
            {
                double channelD = channel / 255.0;
                double result;

                if (tint > 0)
                {
                    // Lighten
                    result = channelD * (1.0 - tint) + (1.0 - (1.0 - tint));
                }
                else
                {
                    // Darken
                    result = channelD * (1.0 + tint);
                }

                // Convert back to 0-255 range and clamp
                int intResult = (int)Math.Round(result * 255);
                return Math.Max(0, Math.Min(255, intResult));
            }

            return Color.FromArgb(
                baseColor.A,
                ApplyChannel(baseColor.R),
                ApplyChannel(baseColor.G),
                ApplyChannel(baseColor.B));
        }
    }
}