// Title: C# – Convert Excel Theme Colors to Explicit RGB with Aspose.Cells
// Description: Loads an Excel workbook, disables theme reliance, iterates every used cell, replaces font, foreground, and background ThemeColor values with their concrete RGB equivalents via Workbook.GetThemeColor, clears ThemeColor references, and saves the file with only explicit RGB formatting.
// Keywords: Aspose.Cells | C# | Excel theme colors | RGB conversion | disable theme | Workbook.GetThemeColor | cell style | explicit color | theme to RGB | convert themed colors
// Common Searches: How to convert Excel theme colors to RGB using Aspose.Cells C# | Aspose.Cells replace theme colors with RGB | Disable Excel theme in .NET workbook | Get concrete RGB from theme color Aspose.Cells | Convert themed cell formatting programmatically
// Developer Intent: The developer wants to load an Excel file, remove all theme‑based color references, and rewrite those colors as fixed RGB values so the workbook no longer depends on a theme.
// Use Cases: Standardize colors for distribution to environments that ignore Excel themes. | Maintain visual consistency when opening workbooks in non‑Microsoft spreadsheet viewers. | Pre‑process bulk Excel files before exporting to PDF or HTML where theme colors may be lost. | Ensure compatibility with older Office versions that lack theme support.
// AI Prompts: Generate C# code with Aspose.Cells that replaces every ThemeColor in a workbook with its RGB value and saves the result. | Explain how Workbook.GetThemeColor works and why clearing ThemeColor properties is required for explicit RGB formatting. | Provide a performance‑optimized version that processes only cells containing ThemeColor values. | Show how to log the number of cells updated during theme‑to‑RGB conversion. | Create a reusable method that accepts a Workbook object and returns a theme‑free workbook.

using System;
using System.Drawing;
using Aspose.Cells;

namespace ThemeConversionExample
{
    // Loads an Excel workbook, disables theme reliance, iterates every used cell, replaces font, foreground, and background ThemeColor values with their concrete RGB equivalents via Workbook.GetThemeColor, clears ThemeColor references, and saves the file with only explicit RGB formatting.
    class Program
    {
        static void Main()
        {
            // Load the workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;

                // Determine the used range
                int maxRow = cells.MaxDataRow;
                int maxCol = cells.MaxDataColumn;

                // Loop through each cell in the used range
                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        Cell cell = cells[row, col];
                        Style style = cell.GetStyle();
                        bool styleChanged = false;

                        // ----- Font Theme Color -----
                        ThemeColor fontTheme = style.Font.ThemeColor;
                        if (fontTheme != null && fontTheme.ColorType != ThemeColorType.StyleColor)
                        {
                            // Get the concrete RGB color from the workbook's theme
                            Color rgbColor = workbook.GetThemeColor(fontTheme.ColorType);
                            // Apply explicit color and clear theme reference
                            style.Font.Color = rgbColor;
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

            // Save the workbook with explicit RGB formatting (replace with desired output path)
            workbook.Save("output.xlsx");
        }
    }
}
