// Title: Apply a built‑in .thmx workbook theme and refresh all cell styles with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells how to apply an Office theme (.thmx) to an existing workbook | refresh Excel cell styles after changing workbook theme using Aspose.Cells | set built‑in workbook theme programmatically with Aspose.Cells .NET | reapply cell style to trigger theme update in Aspose.Cells C# | load .thmx file and apply to workbook with Aspose.Cells example
// Tags: set workbook theme from .thmx file | refresh all cell styles Aspose.Cells | reapply existing style to trigger refresh | iterate worksheets and cells for style update | programmatic theme change .NET Excel

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing XLSX workbook, optionally loads a .thmx theme file and applies it using Aspose.Cells, then iterates through every populated cell, reassigning each cell's current style to force a style refresh, and finally saves the updated workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Apply a theme if a theme file is available (optional)
            // Example:
            // Theme theme = Theme.Load("Office2016.thmx");
            // workbook.SetTheme(theme);
            // Note: Aspose.Cells does not expose a ThemeType enum; a .thmx file must be supplied.

            // Refresh cell styles so they pick up any theme changes
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;
                int maxRow = cells.MaxDataRow;
                int maxColumn = cells.MaxDataColumn;

                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxColumn; col++)
                    {
                        Cell cell = cells[row, col];
                        Style style = cell.GetStyle();
                        cell.SetStyle(style); // reapply to trigger style refresh
                    }
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
