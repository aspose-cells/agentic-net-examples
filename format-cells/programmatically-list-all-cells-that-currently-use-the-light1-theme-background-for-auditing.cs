// Title: How to list Excel cells that use the Light1 theme background color with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# find cells with Light1 theme background color | enumerate Excel cells that have Light1 theme fill using .NET | audit workbook for cells using Light1 theme color in C# Aspose.Cells | retrieve addresses of cells with Light1 background theme in an .xlsx file | list all cells styled with Light1 theme color using Aspose.Cells API
// Tags: Aspose.Cells enumerate cells by theme background | C# list cells with Light1 theme fill | Excel theme color auditing using Aspose.Cells .NET | retrieve cell addresses based on background theme color | iterate worksheets and check cell style theme color

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, iterates through each worksheet and cell, checks each cell's Style.BackgroundThemeColor for the Light1 value, collects matching cell addresses (e.g., "Sheet1!A1"), and prints the list for auditing purposes.
class ListLight1ThemeBackgroundCells
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // List to hold addresses of cells using Light1 theme background
            List<string> light1Cells = new List<string>();

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all cells that have data or formatting
                foreach (Cell cell in sheet.Cells)
                {
                    // Get the style of the current cell
                    Style style = cell.GetStyle();

                    // Check if the background theme color is Light1 (using string comparison for safety)
                    if (style.BackgroundThemeColor.ToString() == "Light1")
                    {
                        // Add the cell address (e.g., "Sheet1!A1") to the list
                        light1Cells.Add($"{sheet.Name}!{cell.Name}");
                    }
                }
            }

            // Output the results
            Console.WriteLine("Cells using Light1 theme background:");
            foreach (string address in light1Cells)
            {
                Console.WriteLine(address);
            }
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
