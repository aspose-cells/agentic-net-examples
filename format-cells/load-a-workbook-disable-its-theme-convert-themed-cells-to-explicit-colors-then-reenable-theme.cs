using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsThemeConversion
{
    class Program
    {
        static void Main()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Verify input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Iterate through all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Iterate through all used cells in the worksheet
                    foreach (Cell cell in sheet.Cells)
                    {
                        // Get the current style of the cell
                        Style style = cell.GetStyle();
                        bool styleChanged = false;

                        // ----- Font Theme Color -----
                        if (style.Font.ThemeColor != null)
                        {
                            // Resolve the actual theme color
                            Color actualColor = workbook.GetThemeColor(style.Font.ThemeColor.ColorType);
                            // Apply explicit color and clear theme reference
                            style.Font.Color = actualColor;
                            style.Font.ThemeColor = null;
                            styleChanged = true;
                        }

                        // ----- Foreground (fill) Theme Color -----
                        if (style.ForegroundThemeColor != null)
                        {
                            Color actualColor = workbook.GetThemeColor(style.ForegroundThemeColor.ColorType);
                            style.ForegroundColor = actualColor;
                            style.ForegroundThemeColor = null;
                            styleChanged = true;
                        }

                        // ----- Background Theme Color -----
                        if (style.BackgroundThemeColor != null)
                        {
                            Color actualColor = workbook.GetThemeColor(style.BackgroundThemeColor.ColorType);
                            style.BackgroundColor = actualColor;
                            style.BackgroundThemeColor = null;
                            styleChanged = true;
                        }

                        // ----- Border Theme Colors -----
                        foreach (BorderType bt in Enum.GetValues(typeof(BorderType)))
                        {
                            Border border = style.Borders[bt];
                            if (border != null && border.ThemeColor != null)
                            {
                                Color actualColor = workbook.GetThemeColor(border.ThemeColor.ColorType);
                                border.Color = actualColor;
                                border.ThemeColor = null;
                                styleChanged = true;
                            }
                        }

                        // Apply the modified style back to the cell if any changes were made
                        if (styleChanged)
                        {
                            cell.SetStyle(style);
                        }
                    }
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}