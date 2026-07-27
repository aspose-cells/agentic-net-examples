using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsThemeConversionDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // -----------------------------------------------------------------
            // Sample data: cells with direct RGB colors (no theme colors set)
            // -----------------------------------------------------------------
            // Cell A1 – direct red foreground
            Cell cellA1 = sheet.Cells["A1"];
            cellA1.PutValue("Red RGB");
            Style styleA1 = cellA1.GetStyle();
            styleA1.ForegroundColor = Color.Red;               // direct RGB
            styleA1.Pattern = BackgroundType.Solid;
            cellA1.SetStyle(styleA1);

            // Cell A2 – direct green background
            Cell cellA2 = sheet.Cells["A2"];
            cellA2.PutValue("Green RGB BG");
            Style styleA2 = cellA2.GetStyle();
            styleA2.BackgroundColor = Color.Green;             // direct RGB
            styleA2.Pattern = BackgroundType.Solid;
            cellA2.SetStyle(styleA2);

            // -----------------------------------------------------------------
            // Conversion: replace direct RGB colors with equivalent theme colors
            // -----------------------------------------------------------------
            // Iterate through all used cells in the worksheet
            foreach (Cell cell in sheet.Cells)
            {
                Style style = cell.GetStyle();

                // ----- Foreground (font) color conversion -----
                // If a theme color is not already assigned and a direct RGB color is present
                if (style.ForegroundThemeColor == null && style.ForegroundColor != Color.Empty)
                {
                    // Try to find a matching theme color
                    ThemeColorType? matchingTheme = FindMatchingThemeColor(workbook, style.ForegroundColor);
                    if (matchingTheme.HasValue)
                    {
                        // Assign the matching theme color (tint = 0.0 for exact match)
                        style.ForegroundThemeColor = new ThemeColor(matchingTheme.Value, 0.0);
                        // Clear the direct RGB value
                        style.ForegroundColor = Color.Empty;
                    }
                }

                // ----- Background color conversion -----
                if (style.BackgroundThemeColor == null && style.BackgroundColor != Color.Empty)
                {
                    ThemeColorType? matchingTheme = FindMatchingThemeColor(workbook, style.BackgroundColor);
                    if (matchingTheme.HasValue)
                    {
                        style.BackgroundThemeColor = new ThemeColor(matchingTheme.Value, 0.0);
                        style.BackgroundColor = Color.Empty;
                    }
                }

                // Apply the possibly modified style back to the cell
                cell.SetStyle(style);
            }

            // Save the workbook
            workbook.Save("ThemeConversionResult.xlsx");
        }

        /// <summary>
        /// Searches the workbook's built‑in theme colors for an exact RGB match.
        /// Returns the matching ThemeColorType if found; otherwise null.
        /// </summary>
        private static ThemeColorType? FindMatchingThemeColor(Workbook workbook, Color rgbColor)
        {
            // Iterate through all defined ThemeColorType values
            foreach (ThemeColorType type in Enum.GetValues(typeof(ThemeColorType)))
            {
                Color themeColor = workbook.GetThemeColor(type);
                if (themeColor.ToArgb() == rgbColor.ToArgb())
                {
                    return type;
                }
            }
            return null; // No exact match found
        }
    }
}