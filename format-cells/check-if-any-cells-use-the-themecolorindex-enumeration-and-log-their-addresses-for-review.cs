using Aspose.Cells;
using System;
using System.Collections.Generic;

class ThemeColorChecker
{
    static void Main()
    {
        // Load an existing workbook (replace the path with your file)
        Workbook workbook = new Workbook("input.xlsx");

        // List to hold addresses of cells that use ThemeColor
        List<string> themeColorCells = new List<string>();

        // Iterate through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            Cells cells = sheet.Cells;

            // Determine the used range to limit iteration
            int maxRow = cells.MaxDataRow;
            int maxCol = cells.MaxDataColumn;

            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = cells[row, col];

                    // Skip null or empty cells
                    if (cell == null || cell.Type == CellValueType.IsNull)
                        continue;

                    // Get the cell's style
                    Style style = cell.GetStyle();

                    bool usesTheme = false;

                    // Font theme color
                    if (style.Font != null && style.Font.ThemeColor != null)
                        usesTheme = true;

                    // Foreground theme color
                    if (style.ForegroundThemeColor != null)
                        usesTheme = true;

                    // Background theme color (if available)
                    if (style.BackgroundThemeColor != null)
                        usesTheme = true;

                    if (usesTheme)
                    {
                        // Record the address in "SheetName!A1" format
                        themeColorCells.Add($"{sheet.Name}!{cell.Name}");
                    }
                }
            }
        }

        // Output the results
        Console.WriteLine("Cells that use ThemeColor:");
        foreach (string address in themeColorCells)
        {
            Console.WriteLine(address);
        }

        // Save the workbook (no modifications made, just to follow lifecycle rule)
        workbook.Save("output.xlsx");
    }
}