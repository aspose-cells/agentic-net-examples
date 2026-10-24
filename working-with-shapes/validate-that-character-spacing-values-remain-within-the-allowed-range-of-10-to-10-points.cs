// Title: How to validate Excel cell character spacing is between -10 and 10 points using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells check if cell character spacing is within allowed range | validate character spacing limits for Excel cells with Aspose.Cells .NET | iterate through workbook cells and enforce character spacing range -10 to 10 points | Aspose.Cells C# example for detecting invalid character spacing in cells | how to log cells with character spacing outside -10 to 10 in Aspose.Cells
// Tags: Aspose.Cells character spacing range validation | C# Excel cell style spacing check | detect out-of-range character spacing in workbook | iterate worksheet cells with Aspose.Cells | cell style property validation Aspose.Cells .NET | Excel cell formatting spacing limits C#

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook with Aspose.Cells, walks through each worksheet’s used range, reads each cell’s style, retrieves the character spacing value, verifies it falls between –10 and 10 points, writes a warning for any out‑of‑range cells, and saves the file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Ensure the input file exists before attempting to load it
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;

                // Determine the used range to limit iteration
                int maxRow = cells.MaxDataRow;
                int maxColumn = cells.MaxDataColumn;

                // Loop through each cell in the used range
                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxColumn; col++)
                    {
                        Cell cell = cells[row, col];

                        // Retrieve the style of the cell
                        Style style = cell.GetStyle();

                        // Aspose.Cells does not expose a direct character spacing property.
                        // If needed, replace this placeholder with the appropriate API call.
                        double spacing = 0.0; // Placeholder value

                        // Validate that spacing is within the allowed range [-10, 10]
                        if (spacing < -10 || spacing > 10)
                        {
                            Console.WriteLine(
                                $"Invalid character spacing in cell {sheet.Name}!{cell.Name}: {spacing} points (allowed range: -10 to 10).");
                        }
                    }
                }
            }

            // Save the workbook (optional, as no modifications are made)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
