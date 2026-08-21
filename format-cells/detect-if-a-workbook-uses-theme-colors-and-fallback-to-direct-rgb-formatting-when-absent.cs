// Title: C# – Detect and Replace Excel Theme Colors with Direct RGB in Aspose.Cells
// Description: Sample code that loads or creates a workbook, checks every cell for font, foreground or background ThemeColor objects, converts each to its actual RGB value via Workbook.GetThemeColor, clears the ThemeColor reference, and saves the file with pure RGB formatting. Demonstrates a safe fallback when themes are unavailable.
// Keywords: Aspose.Cells C# theme color | convert theme to RGB | Workbook.GetThemeColor | replace ThemeColor with RGB | Excel theme fallback | detect theme colors programmatically | cell style conversion | Aspose.Cells example
// Common Searches: how to convert Excel theme colors to RGB using Aspose.Cells | replace theme colors with RGB in C# workbook | detect theme color usage in Aspose.Cells | fallback to direct RGB when theme missing | Aspose.Cells get theme color value
// Developer Intent: Identify any ThemeColor styling in a workbook and substitute it with the equivalent RGB color to ensure theme‑independent formatting.
// Use Cases: Exporting to PDF, images, or formats that ignore Excel themes. | Standardizing styling for downstream processing pipelines. | Sharing workbooks with systems that lack theme support.
// AI Prompts: Generate a C# method that scans an Aspose.Cells workbook and replaces all ThemeColor references with their RGB equivalents. | Show code to log cell addresses that contain theme‑based colors before conversion. | Create a reusable Aspose.Cells utility class for theme‑to‑RGB conversion of fonts, fills, and backgrounds.

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsThemeFallback
{
    // Sample code that loads or creates a workbook, checks every cell for font, foreground or background ThemeColor objects, converts each to its actual RGB value via Workbook.GetThemeColor, clears the ThemeColor reference, and saves the file with pure RGB formatting. Demonstrates a safe fallback when themes are unavailable.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (or load an existing one if needed)
                Workbook workbook = new Workbook(); // creates a new workbook
                Worksheet sheet = workbook.Worksheets[0];

                // Sample: apply a theme color to a cell to demonstrate detection
                Cell sampleCell = sheet.Cells["A1"];
                sampleCell.PutValue("Theme Color Cell");
                Style sampleStyle = sampleCell.GetStyle();

                // Apply a font theme color (Accent1)
                sampleStyle.Font.ThemeColor = new ThemeColor(ThemeColorType.Accent1, 0.0);

                // Set a solid fill pattern
                sampleStyle.Pattern = BackgroundType.Solid;

                // Apply a foreground theme color (Accent2) via direct RGB to avoid setter issue
                // This demonstrates the fallback logic without triggering a NullReferenceException
                Color accent2 = workbook.GetThemeColor(ThemeColorType.Accent2);
                sampleStyle.ForegroundColor = accent2;

                sampleCell.SetStyle(sampleStyle);

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
                            if (cell == null) continue; // Skip unused cells

                            Style style = cell.GetStyle();
                            bool changed = false;

                            // Replace font theme color with actual RGB color
                            if (style.Font.ThemeColor != null)
                            {
                                ThemeColor tc = style.Font.ThemeColor;
                                Color actual = workbook.GetThemeColor(tc.ColorType);
                                style.Font.Color = actual;          // set direct color
                                style.Font.ThemeColor = null;       // remove theme reference
                                changed = true;
                            }

                            // Replace foreground theme color with actual RGB color
                            if (style.ForegroundThemeColor != null)
                            {
                                ThemeColor tc = style.ForegroundThemeColor;
                                Color actual = workbook.GetThemeColor(tc.ColorType);
                                style.ForegroundColor = actual;
                                style.ForegroundThemeColor = null;
                                changed = true;
                            }

                            // Replace background theme color with actual RGB color
                            if (style.BackgroundThemeColor != null)
                            {
                                ThemeColor tc = style.BackgroundThemeColor;
                                Color actual = workbook.GetThemeColor(tc.ColorType);
                                style.BackgroundColor = actual;
                                style.BackgroundThemeColor = null;
                                changed = true;
                            }

                            // Apply the modified style back to the cell if any changes were made
                            if (changed)
                            {
                                cell.SetStyle(style);
                            }
                        }
                    }
                }

                // Save the workbook with direct RGB formatting
                string outputPath = "Output_With_RGB_Colors.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
