// Title: How to list Excel worksheets that contain cells with the Light1 theme background using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# find worksheets with Light1 theme background cells | list Excel sheets that use Light1 background color via Aspose.Cells API | enumerate worksheets containing cells styled with Light1 theme in .NET | C# scan workbook for Light1 themed cell background and get sheet names | retrieve worksheet names where any cell has Light1 background using Aspose.Cells
// Tags: worksheet filtering by cell background theme Aspose.Cells | Light1 theme background detection in Excel cells C# | used range scanning for specific background color Aspose.Cells | retrieve sheet names with themed background cells .NET | Aspose.Cells style-based worksheet enumeration

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook with Aspose.Cells, iterates through each worksheet's used range, checks each cell's style for a non‑none pattern and a BackgroundThemeColor of Light1, collects the names of worksheets that contain at least one such cell, and outputs those worksheet names to the console.
class Program
{
    static void Main()
    {
        // Define input and (optional) output file paths
        string inputPath = "{InputFilePath}";
        string outputPath = "{OutputFilePath}";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the existing workbook
            var workbook = new Workbook(inputPath);

            // List to collect worksheet names that contain cells with Light1 theme background
            var worksheetsWithLight1 = new List<string>();

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                bool containsLight1 = false;

                // Use the used range to limit the scan to cells that actually have data or formatting
                int maxRow = sheet.Cells.MaxDataRow;
                int maxCol = sheet.Cells.MaxDataColumn;

                // Scan each cell in the used range
                for (int row = 0; row <= maxRow && !containsLight1; row++)
                {
                    for (int col = 0; col <= maxCol && !containsLight1; col++)
                    {
                        Cell cell = sheet.Cells[row, col];
                        Style style = cell.GetStyle();

                        // Check if the cell has a background pattern and the background theme color is Light1
                        if (style.Pattern != BackgroundType.None &&
                            style.BackgroundThemeColor.ToString() == "Light1")
                        {
                            containsLight1 = true; // Found at least one matching cell
                        }
                    }
                }

                // If a matching cell was found, add the worksheet name to the result list
                if (containsLight1)
                {
                    worksheetsWithLight1.Add(sheet.Name);
                }
            }

            // Output the names of worksheets that contain Light1 themed background cells
            foreach (string wsName in worksheetsWithLight1)
            {
                Console.WriteLine(wsName);
            }

            // (Optional) Save the workbook if any modifications were made
            // if (!string.IsNullOrWhiteSpace(outputPath))
            // {
            //     workbook.Save(outputPath);
            // }
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
