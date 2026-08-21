// Title: Detect Cells Using ThemeColorIndex (Font, Fill, Background) with Aspose.Cells for .NET
// Description: A C# sample that loads an Excel workbook, scans every used cell, checks the Style object for Font.ThemeColor, ForegroundThemeColor, or BackgroundThemeColor, and writes the addresses of cells that employ the ThemeColorIndex enumeration to the console.
// Keywords: Aspose.Cells ThemeColorIndex detection | C# find theme colored cells | Excel theme color audit .NET | list cells with theme font color | check theme fill color Aspose
// Common Searches: Aspose.Cells how to locate cells with ThemeColorIndex | C# code to list cells using theme colors in Excel | detect theme font or fill color in workbook | search cells for ThemeColor in Aspose.Cells | audit Excel theme colors with .NET
// Developer Intent: Identify every cell whose style references ThemeColorIndex and output its address.
// Use Cases: Verify corporate theme colors are applied consistently before publishing a workbook. | Create a report of theme‑based formatting prior to converting to PDF or image formats. | Ensure no theme colors remain when exporting to formats that do not support them.
// AI Prompts: Generate C# code that collects cell addresses with Font.ThemeColor, ForegroundThemeColor, or BackgroundThemeColor set and stores them in a List<string>. | Optimize the ThemeColor checker to skip rows/columns without formatting and export results to a CSV file. | Explain how to extend the example to capture ThemeColorIndex used in custom styles and include the color type in the output.

using System;
using Aspose.Cells;

namespace ThemeColorChecker
{
    // A C# sample that loads an Excel workbook, scans every used cell, checks the Style object for Font.ThemeColor, ForegroundThemeColor, or BackgroundThemeColor, and writes the addresses of cells that employ the ThemeColorIndex enumeration to the console.
    class Program
    {
        static void Main(string[] args)
        {
            // Load an existing workbook (replace with your file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;
                // Determine the used range
                int maxRow = cells.MaxDataRow;
                int maxCol = cells.MaxDataColumn;

                // Scan each cell within the used range
                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        Cell cell = cells[row, col];
                        // Skip empty cells
                        if (cell == null || cell.Type == CellValueType.IsNull) continue;

                        Style style = cell.GetStyle();

                        // Check for theme color usage in font
                        if (style.Font.ThemeColor != null)
                        {
                            Console.WriteLine($"Cell {cell.Name} uses ThemeColor in Font (Type: {style.Font.ThemeColor.ColorType})");
                        }

                        // Check for theme color usage in foreground/background
                        if (style.ForegroundThemeColor != null)
                        {
                            Console.WriteLine($"Cell {cell.Name} uses ThemeColor in Foreground (Type: {style.ForegroundThemeColor.ColorType})");
                        }

                        if (style.BackgroundThemeColor != null)
                        {
                            Console.WriteLine($"Cell {cell.Name} uses ThemeColor in Background (Type: {style.BackgroundThemeColor.ColorType})");
                        }
                    }
                }
            }

            // Save the workbook (optional, unchanged)
            workbook.Save("output.xlsx");
        }
    }
}
