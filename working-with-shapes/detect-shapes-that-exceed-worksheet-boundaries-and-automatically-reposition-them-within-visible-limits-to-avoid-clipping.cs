// Title: C# code to automatically detect and reposition shapes that exceed worksheet boundaries using Aspose.Cells
// Common Searches: C# Aspose.Cells move shape that is outside the used range of a worksheet | How to adjust shape position to prevent clipping in Excel with Aspose.Cells .NET | Detect shapes extending beyond max row or column and reposition them using Aspose.Cells | Aspose.Cells shape UpperLeftRow UpperLeftColumn correction for out‑of‑bounds shapes in C# | Programmatically ensure Excel shapes stay within visible area with Aspose.Cells
// Tags: Aspose.Cells shape boundary adjustment | C# detect out-of-bounds Excel shapes | reposition Excel shapes programmatically | prevent shape clipping Aspose.Cells | shape coordinate correction .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an Excel workbook, iterates through each worksheet and its shapes, checks each shape's UpperLeftRow/Column and LowerRightRow/Column against the sheet's maximum used row and column, and automatically repositions any shape that lies outside those limits, logging the changes before saving the updated file.
class ShapeBoundaryAdjuster
{
    static void Main()
    {
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        // Verify input file exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file '{inputPath}' not found.");
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

        // Process each worksheet
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Get the last used row/column (zero‑based indices)
            int maxRow = sheet.Cells.MaxRow;
            int maxColumn = sheet.Cells.MaxColumn;

            // Iterate through all shapes on the worksheet
            foreach (Shape shape in sheet.Shapes)
            {
                try
                {
                    int ulRow = shape.UpperLeftRow;
                    int ulCol = shape.UpperLeftColumn;
                    int lrRow = shape.LowerRightRow;
                    int lrCol = shape.LowerRightColumn;

                    bool repositioned = false;

                    // Ensure shape does not start before first row/column
                    if (ulRow < 0)
                    {
                        shape.UpperLeftRow = 0;
                        repositioned = true;
                    }
                    if (ulCol < 0)
                    {
                        shape.UpperLeftColumn = 0;
                        repositioned = true;
                    }

                    // Shift upward if shape extends beyond last visible row
                    if (lrRow > maxRow)
                    {
                        int overflowRows = lrRow - maxRow;
                        int newUlRow = Math.Max(0, ulRow - overflowRows);
                        shape.UpperLeftRow = newUlRow;
                        repositioned = true;
                    }

                    // Shift leftward if shape extends beyond last visible column
                    if (lrCol > maxColumn)
                    {
                        int overflowCols = lrCol - maxColumn;
                        int newUlCol = Math.Max(0, ulCol - overflowCols);
                        shape.UpperLeftColumn = newUlCol;
                        repositioned = true;
                    }

                    // Log moved shapes
                    if (repositioned)
                    {
                        Console.WriteLine($"Shape '{shape.Name}' on sheet '{sheet.Name}' was repositioned within visible limits.");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing shape '{shape.Name}' on sheet '{sheet.Name}': {ex.Message}");
                }
            }
        }

        // Save the modified workbook
        try
        {
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error saving workbook: {ex.Message}");
        }
    }
}
