// Title: C# – Convert Excel Theme Colors to Explicit RGB with Aspose.Cells and Restore Theme
// Description: Loads an Excel workbook, stores the first twelve theme colors, iterates every used cell to replace themed font, fill, background, and border colors with their actual RGB values, clears theme references, re‑applies the saved theme, and saves the result. Demonstrates how to disable a theme, make cell formatting theme‑independent, and then re‑enable the original theme using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# Excel theme conversion | remove theme colors | explicit RGB colors | restore Excel theme | Workbook theme colors | Excel cell style conversion | Aspose.Cells .NET example | theme color to RGB | Excel theme handling
// Common Searches: Aspose.Cells convert themed colors to RGB C# | remove Excel theme formatting programmatically | restore original workbook theme after color change | iterate used cells to replace theme colors Aspose | C# example disabling Excel theme with Aspose.Cells | how to export theme‑independent Excel file
// Developer Intent: Replace every theme‑based formatting in a workbook with concrete RGB values, then re‑apply the original theme to keep the file visually unchanged.
// Use Cases: Generate a theme‑independent copy of a spreadsheet for reliable PDF or image export. | Prepare data for downstream systems that do not support Excel theme colors. | Perform bulk style edits while preserving the ability to restore the original theme later.
// AI Prompts: Create a C# method that extracts the first twelve ThemeColorType values from an Aspose.Cells Workbook and returns a Dictionary<ThemeColorType, Color>. | Provide Aspose.Cells code that converts only border theme colors to explicit RGB values, leaving other style elements untouched. | Explain how to limit cell iteration to the used range of each worksheet for optimal performance during theme‑color conversion.

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using Aspose.Cells;

namespace ThemeConversionDemo
{
    // Loads an Excel workbook, stores the first twelve theme colors, iterates every used cell to replace themed font, fill, background, and border colors with their actual RGB values, clears theme references, re‑applies the saved theme, and saves the result. Demonstrates how to disable a theme, make cell formatting theme‑independent, and then re‑enable the original theme using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // -------------------------------------------------------------
                // Step 1: Store the original theme colors so we can re‑enable later
                // -------------------------------------------------------------
                var originalThemeColors = new Dictionary<ThemeColorType, Color>();
                foreach (ThemeColorType type in Enum.GetValues(typeof(ThemeColorType)))
                {
                    // ThemeColorType has values beyond the 12 theme colors; store only the first 12
                    if ((int)type > 11) continue;
                    originalThemeColors[type] = workbook.GetThemeColor(type);
                }

                // -------------------------------------------------------------
                // Step 2: Convert all themed styles to explicit (non‑themed) colors
                // -------------------------------------------------------------
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Iterate through all used cells in the worksheet
                    foreach (Cell cell in sheet.Cells)
                    {
                        Style style = cell.GetStyle();

                        // ----- Font -----
                        if (style.Font.ThemeColor != null)
                        {
                            Color actualColor = workbook.GetThemeColor(style.Font.ThemeColor.ColorType);
                            style.Font.Color = actualColor;
                            style.Font.ThemeColor = null;
                        }

                        // ----- Foreground (fill) -----
                        if (style.ForegroundThemeColor != null)
                        {
                            Color actualColor = workbook.GetThemeColor(style.ForegroundThemeColor.ColorType);
                            style.ForegroundColor = actualColor;
                            style.ForegroundThemeColor = null;
                        }

                        // ----- Background (fill) -----
                        if (style.BackgroundThemeColor != null)
                        {
                            Color actualColor = workbook.GetThemeColor(style.BackgroundThemeColor.ColorType);
                            style.BackgroundColor = actualColor;
                            style.BackgroundThemeColor = null;
                        }

                        // ----- Borders -----
                        foreach (BorderType borderType in Enum.GetValues(typeof(BorderType)))
                        {
                            Border border = style.Borders[borderType];
                            if (border.ThemeColor != null)
                            {
                                Color actualColor = workbook.GetThemeColor(border.ThemeColor.ColorType);
                                border.Color = actualColor;
                                border.ThemeColor = null;
                            }
                        }

                        // Apply the modified style back to the cell
                        cell.SetStyle(style);
                    }
                }

                // -------------------------------------------------------------
                // Step 3: Re‑enable the original theme (restore original colors)
                // -------------------------------------------------------------
                foreach (var kvp in originalThemeColors)
                {
                    workbook.SetThemeColor(kvp.Key, kvp.Value);
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
