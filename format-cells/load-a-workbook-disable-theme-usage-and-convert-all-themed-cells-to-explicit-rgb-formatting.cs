using System;
using System.Drawing;
using Aspose.Cells;

class ConvertThemedCellsToRgb
{
    static void Main()
    {
        // Load the workbook from a file
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through all worksheets
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Get the used range of cells
            Cells cells = sheet.Cells;
            int maxRow = cells.MaxDataRow;
            int maxCol = cells.MaxDataColumn;

            // Loop through each cell in the used range
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = cells[row, col];
                    // Skip empty cells
                    if (cell == null) continue;

                    // Retrieve the cell style
                    Style style = cell.GetStyle();

                    bool styleChanged = false;

                    // ----- Font Theme Color -----
                    ThemeColor fontTheme = style.Font.ThemeColor;
                    if (fontTheme != null && fontTheme.ColorType != ThemeColorType.StyleColor)
                    {
                        // Get the actual RGB color from the workbook theme
                        Color rgbColor = workbook.GetThemeColor(fontTheme.ColorType);
                        // Apply the color to the font (ignoring tint for simplicity)
                        style.Font.Color = rgbColor;
                        // Clear the theme reference
                        style.Font.ThemeColor = null;
                        styleChanged = true;
                    }

                    // ----- Foreground Theme Color -----
                    ThemeColor fgTheme = style.ForegroundThemeColor;
                    if (fgTheme != null && fgTheme.ColorType != ThemeColorType.StyleColor)
                    {
                        Color rgbColor = workbook.GetThemeColor(fgTheme.ColorType);
                        style.ForegroundColor = rgbColor;
                        style.ForegroundThemeColor = null;
                        styleChanged = true;
                    }

                    // ----- Background Theme Color -----
                    ThemeColor bgTheme = style.BackgroundThemeColor;
                    if (bgTheme != null && bgTheme.ColorType != ThemeColorType.StyleColor)
                    {
                        Color rgbColor = workbook.GetThemeColor(bgTheme.ColorType);
                        style.BackgroundColor = rgbColor;
                        style.BackgroundThemeColor = null;
                        styleChanged = true;
                    }

                    // Apply the modified style back to the cell if any changes were made
                    if (styleChanged)
                    {
                        cell.SetStyle(style);
                    }
                }
            }
        }

        // Save the workbook with explicit RGB formatting
        workbook.Save("output.xlsx");
    }
}