// Title: How to shade alternate rows with the Light1 theme color in an Excel table using Aspose.Cells for .NET
// AI Prompts: Generate C# code that uses Aspose.Cells to fill every other row of a given range with the Light1 theme background while preserving existing formatting. | Demonstrate how to merge a solid Light1 fill style into the current style of cells in alternating rows of a worksheet using Aspose.Cells.
// Common Searches: asp.net apply Light1 theme to alternate rows in Excel with Aspose.Cells | c# banded rows using Excel theme colors via Aspose.Cells | how to set every second row background to Light1 in an Aspose.Cells workbook | Aspose.Cells merge new fill style with existing cell style for row banding
// Tags: Aspose.Cells apply theme color to rows | C# alternate row shading Aspose.Cells | Excel Light1 fill style Aspose.Cells | merge cell style with existing formatting Aspose.Cells | banded rows using Aspose.Cells .NET

using Aspose.Cells;
using System;
using System.Drawing;
using System.IO;

// The program loads an existing workbook, defines a rectangular range (A1:D10), creates a solid LightGray fill style as a fallback, iterates through the rows, and for every second row merges this fill with each cell's current style before saving the modified workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Define the table range (example: A1:D10)
            int startRow = 0;      // Row 1 (zero‑based)
            int endRow = 9;        // Row 10
            int startColumn = 0;   // Column A
            int endColumn = 3;     // Column D

            // Use a fallback color (LightGray) for alternate rows
            Color alternateColor = Color.LightGray;

            // Create a style that uses the fallback color as a solid background
            Style altStyle = workbook.CreateStyle();
            altStyle.ForegroundColor = alternateColor;
            altStyle.Pattern = BackgroundType.Solid;

            // Apply the style to alternate rows (e.g., rows 2,4,6,…)
            for (int row = startRow; row <= endRow; row++)
            {
                // Apply to every second row starting with the second row in the range
                if ((row - startRow) % 2 == 1)
                {
                    for (int col = startColumn; col <= endColumn; col++)
                    {
                        Cell cell = sheet.Cells[row, col];
                        // Preserve other style attributes by merging with the existing style
                        Style current = cell.GetStyle();
                        current.ForegroundColor = altStyle.ForegroundColor;
                        current.Pattern = altStyle.Pattern;
                        cell.SetStyle(current);
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
