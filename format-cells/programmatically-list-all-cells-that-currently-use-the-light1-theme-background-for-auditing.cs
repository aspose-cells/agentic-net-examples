// Title: C# – List cells that use the Light1 (Background1) theme background with Aspose.Cells
// Description: Loads an Excel workbook, walks through every worksheet and each existing cell, reads the cell style, checks if the BackgroundThemeColor equals ThemeColorType.Background1 (Light1), and writes the full address (SheetName!Cell) to the console. No workbook changes are made.
// Keywords: Aspose.Cells | C# | .NET | Light1 theme background | Background1 theme color | audit Excel theme colors | list cells by theme color | theme color detection
// Common Searches: Aspose.Cells find cells with Light1 background | C# list cells using Background1 theme color | how to audit Excel theme colors with Aspose.Cells | enumerate cells by theme background in .NET
// Developer Intent: Retrieve the addresses of every cell that uses the Light1 (Background1) theme background in a workbook.
// Use Cases: Verify corporate theme compliance by locating default Light1 cells. | Generate a report of cells that need style updates before applying a new theme. | Pre‑process a workbook to replace Light1 backgrounds with custom fills.
// AI Prompts: Create C# Aspose.Cells code that returns a List<string> of all cell addresses with a Light1 background theme. | Modify the sample to replace Light1 theme backgrounds with a solid RGB color while keeping other formatting. | Write code that logs worksheet name and cell address for each Background1 themed cell to a text file.

using System;
using Aspose.Cells;

namespace AsposeCellsThemeAudit
{
    // Loads an Excel workbook, walks through every worksheet and each existing cell, reads the cell style, checks if the BackgroundThemeColor equals ThemeColorType.Background1 (Light1), and writes the full address (SheetName!Cell) to the console. No workbook changes are made.
    class Program
    {
        static void Main()
        {
            // Load the workbook to be audited
            Workbook workbook = new Workbook("input.xlsx");

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;

                // Enumerate each cell that actually exists in the worksheet
                foreach (Cell cell in cells)
                {
                    // Retrieve the cell's style
                    Style style = cell.GetStyle();

                    // Get the background theme color (null if not a theme color)
                    ThemeColor bgTheme = style.BackgroundThemeColor;

                    // Check if the background theme color is Light1 (Background1)
                    if (bgTheme != null && bgTheme.ColorType == ThemeColorType.Background1)
                    {
                        // Output the cell address (e.g., "A1")
                        Console.WriteLine($"{sheet.Name}!{cell.Name}");
                    }
                }
            }

            // Optionally, save the workbook if any modifications were made
            // workbook.Save("output.xlsx");
        }
    }
}
