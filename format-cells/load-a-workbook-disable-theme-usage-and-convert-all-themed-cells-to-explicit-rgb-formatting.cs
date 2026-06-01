using System;
using System.Drawing;
using Aspose.Cells;

class ConvertThemeToRgb
{
    static void Main()
    {
        // Load the workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through all worksheets
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Iterate through all used cells
            int maxRow = sheet.Cells.MaxDataRow;
            int maxCol = sheet.Cells.MaxDataColumn;

            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = sheet.Cells[row, col];
                    Style style = cell.GetStyle();
                    bool styleChanged = false;

                    // Convert font theme color to explicit RGB
                    if (style.Font.ThemeColor != null)
                    {
                        ThemeColor tc = style.Font.ThemeColor;
                        Color themeClr = workbook.GetThemeColor(tc.ColorType);
                        style.Font.Color = themeClr;          // Set explicit color
                        style.Font.ThemeColor = null;         // Remove theme reference
                        styleChanged = true;
                    }

                    // Convert foreground theme color to explicit RGB
                    if (style.ForegroundThemeColor != null)
                    {
                        ThemeColor tc = style.ForegroundThemeColor;
                        Color themeClr = workbook.GetThemeColor(tc.ColorType);
                        style.ForegroundColor = themeClr;
                        style.ForegroundThemeColor = null;
                        styleChanged = true;
                    }

                    // Convert background theme color to explicit RGB
                    if (style.BackgroundThemeColor != null)
                    {
                        ThemeColor tc = style.BackgroundThemeColor;
                        Color themeClr = workbook.GetThemeColor(tc.ColorType);
                        style.BackgroundColor = themeClr;
                        style.BackgroundThemeColor = null;
                        styleChanged = true;
                    }

                    // Apply modified style back to the cell
                    if (styleChanged)
                    {
                        cell.SetStyle(style);
                    }
                }
            }
        }

        // Save the workbook with explicit RGB colors
        workbook.Save("output.xlsx");
    }
}