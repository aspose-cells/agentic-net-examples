// Title: How to list all cells that use the Light1 theme background color in an Excel workbook with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code using Aspose.Cells to scan every worksheet in a workbook and return the addresses of cells whose BackgroundThemeColor equals Light1. | Create a reusable method that takes a Workbook object and a ThemeColor enum value, then returns a collection of cell references with that theme background. | Extend the example to group the matching cells by worksheet and optionally export the results to a CSV file.
// Common Searches: aspocells c# find cells with Light1 theme background color in workbook | enumerate Excel cells by theme color using Aspose.Cells .NET | audit Excel file for cells using specific theme background color in C# | list cell addresses with Light1 background theme in Aspose.Cells example | how to get cells with ThemeColor Light1 in Aspose.Cells C#
// Tags: Aspose.Cells enumerate cells by BackgroundThemeColor | C# list Excel cells with Light1 theme color | audit workbook cells based on theme background | retrieve cell addresses for specific theme color Aspose.Cells | scan worksheets for theme background color in .NET

using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.IO;

// The program loads an Excel file with Aspose.Cells, iterates through each worksheet's used range, checks each cell's Style.BackgroundThemeColor for the value Light1, collects the cell addresses (including sheet name), and prints the list to the console.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        Workbook workbook;
        try
        {
            // Load the workbook
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load workbook: {ex.Message}");
            return;
        }

        // List to hold addresses of cells using Light1 theme background
        List<string> light1Cells = new List<string>();

        try
        {
            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Determine the used range to limit iteration
                int maxRow = sheet.Cells.MaxDataRow;
                int maxCol = sheet.Cells.MaxDataColumn;

                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        Cell cell = sheet.Cells[row, col];
                        Style style = cell.GetStyle();

                        // Check if the cell's background theme color is Light1
                        // Use string comparison to avoid enum member issues across versions
                        if (style.BackgroundThemeColor.ToString() == "Light1")
                        {
                            // Record the cell address with its worksheet name
                            light1Cells.Add($"{sheet.Name}!{cell.Name}");
                        }
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error while processing workbook: {ex.Message}");
            return;
        }

        // Output the results
        Console.WriteLine("Cells using Light1 theme background:");
        foreach (string address in light1Cells)
        {
            Console.WriteLine(address);
        }
    }
}
