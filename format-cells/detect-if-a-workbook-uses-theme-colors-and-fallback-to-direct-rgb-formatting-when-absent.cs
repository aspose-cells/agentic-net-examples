// Title: Convert Theme Colors to Direct RGB in Aspose.Cells for .NET
// Description: Demonstrates how to detect theme‑based font, foreground, and background colors in a workbook, retrieve their RGB equivalents with Workbook.GetThemeColor, replace the ThemeColor references, and save the file using only direct RGB formatting.
// Keywords: Aspose.Cells theme to RGB | C# convert theme colors | Workbook.GetThemeColor example | fallback theme colors Aspose | replace ThemeColor with Color | .NET Excel theme color conversion
// Common Searches: how to replace theme colors with RGB in Aspose.Cells | Aspose.Cells get theme color RGB value | convert Excel theme colors to solid colors C# | remove theme color dependency Aspose.Cells | detect and fallback theme colors Aspose
// Developer Intent: Swap any ThemeColor used in cell styles for its concrete RGB Color before saving the workbook.
// Use Cases: Ensure Excel files display correctly in viewers that lack theme support. | Create backward‑compatible spreadsheets for older Office versions. | Standardize visual appearance across workbooks by eliminating theme dependencies.
// AI Prompts: Generate C# code that scans a Workbook and replaces all ThemeColor properties with the corresponding RGB Color using Aspose.Cells. | Explain how Workbook.GetThemeColor can be used to retrieve the base color for a ThemeColorType and apply it to cell styles. | Show a step‑by‑step method to detect theme colors in a workbook and perform a fallback conversion to direct RGB formatting.

using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsThemeFallbackDemo
{
    // Demonstrates how to detect theme‑based font, foreground, and background colors in a workbook, retrieve their RGB equivalents with Workbook.GetThemeColor, replace the ThemeColor references, and save the file using only direct RGB formatting.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (or load an existing one)
                Workbook workbook = new Workbook();

                // Example: add some cells that use theme colors
                Worksheet sheet = workbook.Worksheets[0];

                // Cell A1 – theme font color
                Cell cell1 = sheet.Cells["A1"];
                cell1.PutValue("Theme Font");
                Style style1 = cell1.GetStyle();
                style1.Font.ThemeColor = new ThemeColor(ThemeColorType.Accent1, 0.0);
                cell1.SetStyle(style1);

                // Cell A2 – theme foreground (fill) color
                Cell cell2 = sheet.Cells["A2"];
                cell2.PutValue("Theme Foreground");
                Style style2 = cell2.GetStyle();
                style2.Pattern = BackgroundType.Solid;                         // Set pattern before theme color
                style2.ForegroundThemeColor = new ThemeColor(ThemeColorType.Accent2, 0.0);
                cell2.SetStyle(style2);

                // Cell A3 – theme background color
                Cell cell3 = sheet.Cells["A3"];
                cell3.PutValue("Theme Background");
                Style style3 = cell3.GetStyle();
                style3.Pattern = BackgroundType.Solid;                         // Set pattern before theme color
                style3.BackgroundThemeColor = new ThemeColor(ThemeColorType.Accent3, 0.0);
                cell3.SetStyle(style3);

                // Iterate through all worksheets and cells to replace theme colors with direct RGB colors
                foreach (Worksheet ws in workbook.Worksheets)
                {
                    Cells cells = ws.Cells;
                    int maxRow = cells.MaxDataRow;
                    int maxCol = cells.MaxDataColumn;

                    for (int row = 0; row <= maxRow; row++)
                    {
                        for (int col = 0; col <= maxCol; col++)
                        {
                            Cell cell = cells[row, col];
                            Style style = cell.GetStyle();
                            bool styleChanged = false;

                            // Replace font theme color with concrete RGB color
                            if (style.Font.ThemeColor != null)
                            {
                                ThemeColor tc = style.Font.ThemeColor;
                                Color themeBase = workbook.GetThemeColor(tc.ColorType);
                                style.Font.Color = themeBase;          // Direct RGB color
                                style.Font.ThemeColor = null;          // Remove theme reference
                                styleChanged = true;
                            }

                            // Replace foreground theme color with concrete RGB color
                            if (style.ForegroundThemeColor != null)
                            {
                                ThemeColor tc = style.ForegroundThemeColor;
                                Color themeBase = workbook.GetThemeColor(tc.ColorType);
                                style.ForegroundColor = themeBase;
                                style.ForegroundThemeColor = null;
                                styleChanged = true;
                            }

                            // Replace background theme color with concrete RGB color
                            if (style.BackgroundThemeColor != null)
                            {
                                ThemeColor tc = style.BackgroundThemeColor;
                                Color themeBase = workbook.GetThemeColor(tc.ColorType);
                                style.BackgroundColor = themeBase;
                                style.BackgroundThemeColor = null;
                                styleChanged = true;
                            }

                            // Apply the modified style back to the cell
                            if (styleChanged)
                            {
                                cell.SetStyle(style);
                            }
                        }
                    }
                }

                // Save the workbook with direct RGB formatting
                workbook.Save("ThemeFallbackResult.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
