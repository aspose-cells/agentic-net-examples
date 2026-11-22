// Title: C# example: List Excel worksheets that contain cells with the Light1 theme background using Aspose.Cells
// Common Searches: how to retrieve worksheet names that have Light1 theme background cells in Aspose.Cells C# | Aspose.Cells code to find Excel sheets with cells using the Light1 theme color | search an Excel workbook for cells with a specific theme background color in .NET | enumerate worksheets based on cell style theme using Aspose.Cells for .NET
// Tags: list worksheets by Light1 theme background Aspose.Cells | detect Light1 theme background cells C# | filter workbook sheets by cell background theme Aspose.Cells | scan used range for specific theme color .NET | enumerate sheets containing particular cell style theme Aspose.Cells

using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.IO;

// The sample loads an Excel file with Aspose.Cells, iterates through each worksheet, scans the used cell range, checks each cell's BackgroundThemeColor for the enum value "Light1", collects the names of worksheets that contain at least one such cell, and outputs the list.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists before attempting to load it
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: File not found – {inputPath}");
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
            Console.WriteLine($"Error loading workbook: {ex.Message}");
            return;
        }

        // List to hold worksheet names that contain Light1 theme background cells
        List<string> sheetsWithLight1 = new List<string>();

        try
        {
            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;
                bool found = false;

                // Determine the used range to limit the search
                int maxRow = cells.MaxDataRow;
                int maxCol = cells.MaxDataColumn;

                // Scan cells within the used range
                for (int row = 0; row <= maxRow && !found; row++)
                {
                    for (int col = 0; col <= maxCol && !found; col++)
                    {
                        Cell cell = cells[row, col];
                        Style style = cell.GetStyle();

                        // Compare the enum name as string to avoid version‑specific enum members
                        if (style.BackgroundThemeColor.ToString() == "Light1")
                        {
                            found = true; // No need to continue scanning this sheet
                        }
                    }
                }

                // If a Light1 background cell was found, add the sheet name to the list
                if (found)
                {
                    sheetsWithLight1.Add(sheet.Name);
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error processing workbook: {ex.Message}");
            return;
        }

        // Output the result
        Console.WriteLine("Worksheets containing Light1 theme background:");
        foreach (string name in sheetsWithLight1)
        {
            Console.WriteLine(name);
        }
    }
}
