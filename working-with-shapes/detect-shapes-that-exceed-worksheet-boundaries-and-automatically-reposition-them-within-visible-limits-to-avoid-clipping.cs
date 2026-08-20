// Title: Auto‑Adjust Out‑Of‑Bounds Shapes in Aspose.Cells (C#)
// Description: C# example that creates a workbook, adds a shape placed beyond the last row/column, detects any shape whose UpperLeftRow or UpperLeftColumn exceeds MaxRow/MaxColumn, and repositions it inside the visible area using MoveToRange before saving.
// Keywords: Aspose.Cells | C# | shape boundary detection | out of bounds shape | move shape to worksheet limits | MaxRow MaxColumn | MoveToRange | Excel automation | adjust shape position | prevent shape clipping
// Common Searches: Aspose.Cells detect shape outside worksheet | C# move Excel shape back into visible area | adjust out‑of‑bounds shapes Aspose.Cells | prevent shape clipping in generated Excel file | reposition shapes beyond last row column Aspose
// Developer Intent: Automatically find shapes that lie outside the worksheet’s usable range and relocate them so they remain visible in the final Excel file.
// Use Cases: Guarantee that programmatically added charts, images, or diagrams are not hidden when the sheet size changes. | Correct imported drawings that were positioned beyond the sheet’s maximum rows or columns before saving. | Create templates where every shape must stay inside the printable or viewable area of the worksheet.
// AI Prompts: Write a C# method that scans all shapes in a worksheet and moves any shape whose UpperLeftRow or UpperLeftColumn is greater than the sheet’s MaxRow/MaxColumn using Aspose.Cells. | Show how to log original and new coordinates for each adjusted shape and optionally resize it to fit within target cells. | Explain how to handle shapes placed on merged cells when repositioning them within worksheet boundaries.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// C# example that creates a workbook, adds a shape placed beyond the last row/column, detects any shape whose UpperLeftRow or UpperLeftColumn exceeds MaxRow/MaxColumn, and repositions it inside the visible area using MoveToRange before saving.
class ShapeBoundaryAdjuster
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a rectangle shape placed far outside typical worksheet limits
            int outOfBoundsRow = 5000;
            int outOfBoundsColumn = 5000;
            int shapeHeight = 100; // points
            int shapeWidth = 200;  // points

            Shape outOfBoundsShape = sheet.Shapes.AddRectangle(
                outOfBoundsRow, outOfBoundsColumn, 0, 0, shapeHeight, shapeWidth);
            outOfBoundsShape.Name = "OutOfBoundsRect";

            // -----------------------------------------------------------------
            // Detect and reposition shapes that exceed worksheet boundaries
            // -----------------------------------------------------------------

            // Maximum allowed row and column indices (zero‑based)
            int maxRow = sheet.Cells.MaxRow;
            int maxColumn = sheet.Cells.MaxColumn;

            // Iterate through all shapes in the worksheet
            for (int i = 0; i < sheet.Shapes.Count; i++)
            {
                Shape shape = sheet.Shapes[i];

                // Current position of the shape
                int shapeRow = shape.UpperLeftRow;
                int shapeColumn = shape.UpperLeftColumn;

                bool needsReposition = false;

                // If the shape starts beyond the last row, move it to the last permissible row
                if (shapeRow > maxRow)
                {
                    shapeRow = maxRow;
                    needsReposition = true;
                }

                // If the shape starts beyond the last column, move it to the last permissible column
                if (shapeColumn > maxColumn)
                {
                    shapeColumn = maxColumn;
                    needsReposition = true;
                }

                if (needsReposition)
                {
                    // Reposition the shape within visible limits (offsets set to 0)
                    shape.MoveToRange(shapeRow, shapeColumn, 0, 0);
                    Console.WriteLine($"Shape '{shape.Name}' repositioned to Row={shapeRow}, Column={shapeColumn}");
                }
            }

            // Ensure the output directory exists
            string outputPath = "AdjustedShapes.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook with the adjusted shapes
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved as '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
