// Title: Convert Excel Theme Colors to RGB and Restore the Original Theme with Aspose.Cells for .NET
// Description: Loads an Excel workbook, captures the first twelve theme colors, walks every used cell across all worksheets, replaces Font, Fill, and Border theme colors with their exact RGB values, clears ThemeColor references, then re‑applies the saved theme before saving the file.
// Keywords: Aspose.Cells | C# | Excel theme colors | convert theme to RGB | disable workbook theme | restore workbook theme | explicit cell colors | flatten Excel colors | theme color conversion .NET | theme‑based formatting
// Common Searches: Aspose.Cells replace theme colors with RGB | How to disable Excel theme using Aspose.Cells C# | Convert themed cells to explicit colors Aspose | Restore original theme after color conversion Aspose.Cells | Flatten workbook colors for PDF export
// Developer Intent: Replace all theme‑based formatting with concrete RGB values, then re‑apply the original workbook theme.
// Use Cases: Create a theme‑free copy for PDF or image export | Prepare a spreadsheet for systems that ignore Excel themes | Archive a workbook with fixed colors while preserving the original theme | Generate a printable version with consistent colors across platforms
// AI Prompts: Provide C# Aspose.Cells code that saves the current theme, converts every themed font, fill, and border color to its RGB equivalent, clears ThemeColor references, and finally restores the saved theme. | Explain how to iterate through all used cells in a workbook and replace ThemeColor objects with explicit Color values while keeping a backup of the theme for later re‑application.

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using Aspose.Cells;

namespace ThemeConversionDemo
{
    // Loads an Excel workbook, captures the first twelve theme colors, walks every used cell across all worksheets, replaces Font, Fill, and Border theme colors with their exact RGB values, clears ThemeColor references, then re‑applies the saved theme before saving the file.
    class Program
    {
        static void Main(string[] args)
        {
            // Paths for input and output workbooks
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Preserve original theme colors so we can re‑enable the theme later
                var originalThemeColors = new Dictionary<ThemeColorType, Color>();
                foreach (ThemeColorType type in Enum.GetValues(typeof(ThemeColorType)))
                {
                    // Only the first 12 types are valid theme colors
                    if ((int)type > 11) break;
                    originalThemeColors[type] = workbook.GetThemeColor(type);
                }

                // Iterate through all worksheets and cells
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Determine the used range
                    int maxRow = sheet.Cells.MaxDataRow;
                    int maxCol = sheet.Cells.MaxDataColumn;

                    for (int row = 0; row <= maxRow; row++)
                    {
                        for (int col = 0; col <= maxCol; col++)
                        {
                            Cell cell = sheet.Cells[row, col];
                            // Skip empty cells
                            if (cell == null) continue;

                            Style style = cell.GetStyle();
                            bool styleChanged = false;

                            // Convert Font theme color to explicit color
                            if (style.Font.ThemeColor != null)
                            {
                                ThemeColor tc = style.Font.ThemeColor;
                                Color explicitColor = workbook.GetThemeColor(tc.ColorType);
                                style.Font.Color = explicitColor;
                                style.Font.ThemeColor = null;
                                styleChanged = true;
                            }

                            // Convert Foreground theme color to explicit color
                            if (style.ForegroundThemeColor != null)
                            {
                                ThemeColor tc = style.ForegroundThemeColor;
                                Color explicitColor = workbook.GetThemeColor(tc.ColorType);
                                style.ForegroundColor = explicitColor;
                                style.ForegroundThemeColor = null;
                                styleChanged = true;
                            }

                            // Convert Background theme color to explicit color
                            if (style.BackgroundThemeColor != null)
                            {
                                ThemeColor tc = style.BackgroundThemeColor;
                                Color explicitColor = workbook.GetThemeColor(tc.ColorType);
                                style.BackgroundColor = explicitColor;
                                style.BackgroundThemeColor = null;
                                styleChanged = true;
                            }

                            // Convert border theme colors to explicit colors
                            foreach (BorderType bt in Enum.GetValues(typeof(BorderType)))
                            {
                                Border border = style.Borders[bt];
                                if (border != null && border.ThemeColor != null)
                                {
                                    ThemeColor tc = border.ThemeColor;
                                    Color explicitColor = workbook.GetThemeColor(tc.ColorType);
                                    border.Color = explicitColor;
                                    border.ThemeColor = null;
                                    styleChanged = true;
                                }
                            }

                            // Apply the modified style back to the cell
                            if (styleChanged)
                            {
                                cell.SetStyle(style);
                            }
                        }
                    }
                }

                // Re‑enable the original theme by restoring its colors
                foreach (var kvp in originalThemeColors)
                {
                    workbook.SetThemeColor(kvp.Key, kvp.Value);
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
