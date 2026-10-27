// Title: Replace the Light2 theme background with a patterned fill while preserving theme colors using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells change Light2 theme background to a Gray50 pattern | How to apply a patterned fill to Excel cells without losing theme colors in .NET | Programmatically replace solid cell fill with a pattern using Aspose.Cells | Iterate through a workbook and set cell background pattern while keeping theme references | Set cell style pattern BackgroundType.Gray50 with Aspose.Cells C# example
// Tags: Aspose.Cells replace theme background pattern | C# set cell background pattern Gray50 | Excel Light2 theme patterned fill Aspose | modify cell style pattern programmatically | preserve theme colors while changing fill

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an existing workbook, loops through all used cells, and for cells that have no pattern (solid fill) changes the style to a Gray50 patterned fill, sets a light gray foreground and white background, then saves the workbook, ensuring the Light2 theme colors remain intact.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Verify input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Iterate through all worksheets and cells
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    Cells cells = sheet.Cells;
                    // Get the maximum used row and column to limit the loop
                    int maxRow = cells.MaxDataRow;
                    int maxCol = cells.MaxDataColumn;

                    for (int row = 0; row <= maxRow; row++)
                    {
                        for (int col = 0; col <= maxCol; col++)
                        {
                            Cell cell = cells[row, col];
                            Style style = cell.GetStyle();

                            // Check for solid fill (no pattern)
                            if (style.Pattern == BackgroundType.None)
                            {
                                // Replace solid fill with a patterned fill
                                style.Pattern = BackgroundType.Gray50; // Example pattern
                                style.ForegroundColor = Color.LightGray; // Preserve a light color for foreground
                                style.BackgroundColor = Color.White;    // Optional background color

                                // Apply the modified style back to the cell
                                cell.SetStyle(style);
                            }
                        }
                    }
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
