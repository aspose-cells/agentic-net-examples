// Title: C# – List All Cells Using the Light1 (Background1) Theme Background with Aspose.Cells
// Description: Loads an Excel workbook, scans each worksheet’s used range, checks each cell’s Style.BackgroundThemeColor for ThemeColorType.Background1 (Light1), collects the full addresses (e.g., Sheet1!A1), prints the list, and saves the file unchanged.
// Keywords: Aspose.Cells | C# | .NET | Light1 theme background | Background1 | ThemeColorType.Background1 | Excel cell style audit | list cells by theme color | cell address enumeration | bulk cell scanning | Excel compliance check
// Common Searches: Aspose.Cells find cells with Light1 background | C# list cells using Background1 theme color | audit Excel workbook for Light1 theme background | retrieve cell addresses by theme color Aspose.Cells | enumerate cells with ThemeColorType.Background1 in .NET
// Developer Intent: Identify and enumerate every cell that currently uses the Light1 (Background1) theme background.
// Use Cases: Generate a report of cells that need background‑theme updates before applying a corporate style. | Validate workbook compliance with design guidelines that forbid the Light1 background. | Create an audit log of theme‑background usage for migration or regulatory purposes.
// AI Prompts: Write C# code with Aspose.Cells that exports the addresses of Light1‑themed cells to a CSV file. | Extend the example to also capture cells using the Dark1 theme background and group results by theme type. | Describe performance‑optimisation strategies for scanning large workbooks for specific theme colors using Aspose.Cells.

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsThemeAudit
{
    // Loads an Excel workbook, scans each worksheet’s used range, checks each cell’s Style.BackgroundThemeColor for ThemeColorType.Background1 (Light1), collects the full addresses (e.g., Sheet1!A1), prints the list, and saves the file unchanged.
    class Program
    {
        static void Main()
        {
            // Load the workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // List to hold addresses of cells using Light1 (Background1) theme background
            List<string> cellsUsingLight1 = new List<string>();

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;

                // Determine the used range to avoid scanning empty cells
                int maxRow = cells.MaxDataRow;
                int maxCol = cells.MaxDataColumn;

                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        // Get the cell; may be null if never accessed
                        Cell cell = cells[row, col];
                        if (cell == null) continue;

                        // Retrieve the cell's style
                        Style style = cell.GetStyle();

                        // Check if the background theme color is set and matches Background1 (Light1)
                        ThemeColor bgTheme = style.BackgroundThemeColor;
                        if (bgTheme != null && bgTheme.ColorType == ThemeColorType.Background1)
                        {
                            // Record the cell address (e.g., "Sheet1!A1")
                            cellsUsingLight1.Add($"{sheet.Name}!{cell.Name}");
                        }
                    }
                }
            }

            // Output the results
            Console.WriteLine("Cells using Light1 (Background1) theme background:");
            foreach (string address in cellsUsingLight1)
            {
                Console.WriteLine(address);
            }

            // Save the workbook (unchanged) – required by lifecycle rule
            workbook.Save("output.xlsx");
        }
    }
}
