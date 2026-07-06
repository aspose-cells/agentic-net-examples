using System;
using Aspose.Cells;

namespace ThemeColorChecker
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your file path) or create a new one if the file does not exist.
            Workbook workbook;
            string inputPath = "input.xlsx";
            if (System.IO.File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook(); // creates a new workbook with one worksheet
            }

            // Iterate through all worksheets.
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all cells that are part of the used range.
                foreach (Cell cell in sheet.Cells)
                {
                    // Retrieve the cell's style.
                    Style style = cell.GetStyle();

                    // Check for any theme color usage in the style.
                    bool usesThemeColor =
                        style.ForegroundThemeColor != null ||
                        style.BackgroundThemeColor != null ||
                        (style.Font != null && style.Font.ThemeColor != null);

                    if (usesThemeColor)
                    {
                        // Log the address of the cell that uses a theme color.
                        Console.WriteLine($"Cell {cell.Name} in worksheet \"{sheet.Name}\" uses a ThemeColor.");
                    }
                }
            }

            // Save the workbook (optional – here we just save a copy).
            workbook.Save("output.xlsx");
        }
    }
}