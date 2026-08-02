// Title: C# – List Excel worksheets with Light1 (Background1) theme background using Aspose.Cells
// Description: Loads a workbook, scans each sheet’s used range, checks Style.BackgroundThemeColor for ThemeColorType.Background1, gathers the names of worksheets that contain at least one matching cell, outputs the list, and saves the file unchanged.
// Keywords: Aspose.Cells C# | list worksheets by theme color | Light1 background detection | Background1 theme color | ThemeColorType.Background1 | Excel cell style inspection | worksheet enumeration Aspose | Excel automation theme colors | detect theme background in cells | C# Excel theme color lookup
// Common Searches: C# Aspose.Cells find sheets with Light1 background | how to list worksheets that use Background1 theme color | detect theme background color in Excel using Aspose | enumerate worksheets containing specific theme color .NET | Aspose.Cells code to get sheet names with ThemeColorType.Background1
// Developer Intent: Retrieve the names of all worksheets that contain at least one cell formatted with the Light1 (Background1) theme background.
// Use Cases: Audit a workbook for default theme colors before applying a corporate style guide. | Generate a checklist of sheets that need color‑scheme updates in a large Excel report. | Create a maintenance script that targets only those worksheets for background‑color replacement.
// AI Prompts: Provide a method that returns a List<string> of worksheet names where any cell uses ThemeColorType.Background1. | Extend the example to also capture the cell addresses (e.g., A1, B2) that have the Light1 background. | Explain how to replace all Light1 background cells with a custom RGB color on the identified worksheets.

using Aspose.Cells;
using System;
using System.Collections.Generic;

// Loads a workbook, scans each sheet’s used range, checks Style.BackgroundThemeColor for ThemeColorType.Background1, gathers the names of worksheets that contain at least one matching cell, outputs the list, and saves the file unchanged.
class ListSheetsWithLight1Background
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Theme color type representing Light1 background (Background1)
        ThemeColorType light1Type = ThemeColorType.Background1;

        // List to hold worksheet names that contain at least one cell with Light1 background
        List<string> sheetsWithLight1 = new List<string>();

        // Iterate through each worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            bool found = false;

            // Determine the used range of the sheet
            int maxRow = sheet.Cells.MaxDataRow;
            int maxCol = sheet.Cells.MaxDataColumn;

            // Skip empty sheets
            if (maxRow < 0 || maxCol < 0)
                continue;

            // Scan cells within the used range
            for (int row = 0; row <= maxRow && !found; row++)
            {
                for (int col = 0; col <= maxCol && !found; col++)
                {
                    // Get the style of the current cell
                    Style style = sheet.Cells[row, col].GetStyle();

                    // Retrieve the background theme color (if any)
                    ThemeColor bgTheme = style.BackgroundThemeColor;

                    // Check if the background theme color matches Light1 (Background1)
                    if (bgTheme != null && bgTheme.ColorType == light1Type)
                    {
                        found = true; // Stop scanning this sheet
                    }
                }
            }

            // If a matching cell was found, add the worksheet name to the list
            if (found)
                sheetsWithLight1.Add(sheet.Name);
        }

        // Output the result
        Console.WriteLine("Worksheets containing cells with Light1 (Background1) theme background:");
        foreach (string name in sheetsWithLight1)
        {
            Console.WriteLine(name);
        }

        // Save the workbook (no modifications made, but required by lifecycle rules)
        workbook.Save("output.xlsx");
    }
}
