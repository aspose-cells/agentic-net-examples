// Title: List worksheet names that contain cells with the Light1 theme background using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that returns every worksheet name where at least one cell has BackgroundThemeColor set to Light1. | Create a function that scans the used range of each sheet in a workbook and collects sheet names that contain any Light1‑themed background cells. | Extend the example to also output the cell addresses of Light1 background cells while still providing the list of matching worksheets.
// Common Searches: asp.net find Excel worksheets with Light1 theme background using Aspose.Cells | c# Aspose.Cells list sheets that contain cells colored Light1 | how to detect Light1 background color in cells across all worksheets with Aspose.Cells | retrieve worksheet names having Light1 theme fill in a .xlsx file using C#
// Tags: enumerate worksheets by cell background theme Aspose.Cells | search used range for Light1 theme color C# | retrieve sheet names with specific theme background Aspose.Cells | detect Light1 background cells in Excel workbook .NET | filter worksheets based on cell style theme Aspose

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The program loads an Excel workbook, iterates through each worksheet's used range, checks each cell's BackgroundThemeColor for the value Light1, collects the names of worksheets that contain at least one such cell, and prints the resulting list.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // List to hold names of worksheets that contain cells with Light1 theme background
            List<string> sheetsWithLight1 = new List<string>();

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                bool found = false;

                // Get the maximum used row and column to limit the search area
                int maxRow = sheet.Cells.MaxDataRow;
                int maxColumn = sheet.Cells.MaxDataColumn;

                // If the sheet is empty, skip it
                if (maxRow < 0 || maxColumn < 0)
                    continue;

                // Scan cells within the used range
                for (int row = 0; row <= maxRow && !found; row++)
                {
                    for (int col = 0; col <= maxColumn && !found; col++)
                    {
                        try
                        {
                            Cell cell = sheet.Cells[row, col];
                            Style style = cell.GetStyle();

                            // Compare the background theme color name with "Light1"
                            if (style.BackgroundThemeColor.ToString() == "Light1")
                            {
                                found = true;
                            }
                        }
                        catch (Exception exCell)
                        {
                            // Log cell-level errors but continue processing
                            Console.WriteLine($"Warning: Unable to process cell [{row}, {col}] in sheet \"{sheet.Name}\": {exCell.Message}");
                        }
                    }
                }

                // If a matching cell was found, add the worksheet name to the result list
                if (found)
                {
                    sheetsWithLight1.Add(sheet.Name);
                }
            }

            // Output the worksheet names
            Console.WriteLine("Worksheets containing cells with Light1 theme background:");
            foreach (string name in sheetsWithLight1)
            {
                Console.WriteLine(name);
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
