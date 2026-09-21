// Title: Replace explicit RGB font and fill colors with matching Excel theme colors using Aspose.Cells for .NET
// AI Prompts: Generate C# code that scans every cell in an Excel workbook with Aspose.Cells, detects direct RGB font or background colors, and substitutes them with the corresponding ThemeColorType. | Write a helper method for Aspose.Cells that receives a System.Drawing.Color, finds the matching ThemeColorType in the workbook's theme palette, and returns the theme color for style updates.
// Common Searches: how to convert explicit RGB cell colors to Excel theme colors with Aspose.Cells in C# | Aspose.Cells replace font color RGB with theme color programmatically | C# iterate through workbook cells and map fill color to theme palette using Aspose | find matching ThemeColorType for a given RGB value in Aspose.Cells | convert workbook background colors to theme colors Aspose.Cells .NET
// Tags: rgb-to-theme color conversion Aspose.Cells | map cell fill color to workbook theme palette C# | detect and replace explicit RGB styles Aspose.Cells | theme color lookup from System.Drawing.Color | bulk cell style update Excel workbook .NET

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;

// The example loads an Excel workbook, walks through every cell in each worksheet, checks the font and fill colors for direct RGB values, and when a color exactly matches a theme color it replaces the RGB value with the appropriate ThemeColorType before saving the file.
class ConvertRgbToThemeColors
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Ensure the input file exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook (lifecycle rule: load)
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Determine the used range limits
                int maxRow = sheet.Cells.MaxDataRow;
                int maxCol = sheet.Cells.MaxDataColumn;

                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        Cell cell = sheet.Cells[row, col];
                        if (cell == null) continue; // Skip null cells

                        // Retrieve the current style
                        Style style = cell.GetStyle();

                        // ----- Process Font Color -----
                        Color fontColor = style.Font.Color;
                        if (!fontColor.IsEmpty)
                        {
                            ThemeColorType? matchingTheme = FindMatchingThemeColor(workbook, fontColor);
                            if (matchingTheme.HasValue)
                            {
                                // Replace with the corresponding theme color
                                style.Font.Color = workbook.GetThemeColor(matchingTheme.Value);
                            }
                        }

                        // ----- Process Cell Background (Foreground) Color -----
                        Color bgColor = style.ForegroundColor;
                        if (!bgColor.IsEmpty)
                        {
                            ThemeColorType? matchingTheme = FindMatchingThemeColor(workbook, bgColor);
                            if (matchingTheme.HasValue)
                            {
                                // Replace with the corresponding theme color
                                style.ForegroundColor = workbook.GetThemeColor(matchingTheme.Value);
                            }
                        }

                        // Apply the possibly modified style back to the cell
                        cell.SetStyle(style);
                    }
                }
            }

            // Save the modified workbook (lifecycle rule: save)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Helper method to locate a ThemeColorType whose actual RGB matches the supplied color
    private static ThemeColorType? FindMatchingThemeColor(Workbook workbook, Color rgbColor)
    {
        foreach (ThemeColorType theme in Enum.GetValues(typeof(ThemeColorType)))
        {
            Color themeRgb = workbook.GetThemeColor(theme);
            if (themeRgb.ToArgb() == rgbColor.ToArgb())
            {
                return theme;
            }
        }
        return null; // No matching theme color found
    }
}
