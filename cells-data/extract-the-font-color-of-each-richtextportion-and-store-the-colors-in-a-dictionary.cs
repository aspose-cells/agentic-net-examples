// Title: How to extract the font color of each Excel cell with Aspose.Cells for .NET and store it in a Sheet!Cell dictionary
// Common Searches: Aspose.Cells .NET get font color of every cell in a workbook | store Excel cell font colors in a C# dictionary using Aspose.Cells | retrieve Font.Color for all cells across worksheets with Aspose.Cells | iterate through cells and read cell style font color Aspose.Cells C# | dictionary key format SheetName!CellAddress for cell colors Aspose.Cells
// Tags: Aspose.Cells cell font color extraction | C# dictionary mapping sheet cell to Color | iterate worksheets and cells Aspose.Cells | read cell style Font.Color | Excel cell color extraction .NET

using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;

// The example loads an Excel workbook with Aspose.Cells, loops through each worksheet and its used cells, obtains the Font.Color from each cell's Style, and saves the color in a Dictionary keyed by "SheetName!CellName" for easy lookup.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Ensure the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Dictionary to store font colors of each cell (or portion if available)
            // Key format: SheetName!CellName
            Dictionary<string, Color> cellColors = new Dictionary<string, Color>();

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all used cells in the worksheet
                foreach (Cell cell in sheet.Cells)
                {
                    // Process only cells that contain a string value
                    if (cell.Value != null && cell.Value is string)
                    {
                        // Retrieve the cell's style
                        Style style = cell.GetStyle();

                        // Extract the font color (if set)
                        Color fontColor = style.Font.Color;

                        // Build a unique key for the dictionary
                        string key = $"{sheet.Name}!{cell.Name}";

                        // Store the color in the dictionary
                        cellColors[key] = fontColor;
                    }
                }
            }

            // Example usage: print the collected colors
            foreach (KeyValuePair<string, Color> entry in cellColors)
            {
                Console.WriteLine($"{entry.Key}: {entry.Value}");
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors during processing
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
