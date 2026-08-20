// Title: Set a shape’s absolute position using column width & row height in Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, set custom column width and row height, convert those dimensions to pixels, add a rectangle shape, move it to a target cell, apply pixel offsets for absolute placement, verify the anchored cell, optionally align the shape’s top‑right corner to another cell, and save the file.
// Keywords: Aspose.Cells shape positioning | pixel offset column width | row height to pixels | MoveToRange shape | shape alignment Aspose.Cells | validate shape cell attachment | C# Aspose.Cells example
// Common Searches: Aspose.Cells set shape absolute position | convert column width to pixels Aspose.Cells | move shape to specific cell .NET | align shape corners in Excel with Aspose.Cells | check shape row and column after MoveToRange
// Developer Intent: Place a rectangle shape at a designated cell using pixel offsets derived from column width and row height, then confirm that the shape is anchored to the intended cell.
// Use Cases: Insert a logo into a header cell with exact left/top offsets for consistent branding across generated reports. | Attach a comment shape to the top‑right corner of a data cell so it moves correctly when rows or columns are resized. | Automate verification of shape anchoring after bulk worksheet processing to ensure layout integrity before publishing.
// AI Prompts: Generate C# code with Aspose.Cells that places a shape at cell B5 with a 10‑pixel left offset and a 5‑pixel top offset, then checks UpperLeftRow and UpperLeftColumn. | Show how to retrieve a column’s width in pixels and align a shape’s right edge with column C while keeping the shape attached to row 8. | Provide an example of error handling that logs a warning if a shape’s actual cell position differs from the expected row and column after calling MoveToRange.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapePositionDemo
{
    // Demonstrates how to create a workbook, set custom column width and row height, convert those dimensions to pixels, add a rectangle shape, move it to a target cell, apply pixel offsets for absolute placement, verify the anchored cell, optionally align the shape’s top‑right corner to another cell, and save the file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Define target cell (row and column) where the shape's upper‑left corner should be placed
                int targetRow = 5;      // zero‑based index (row 6 in Excel)
                int targetColumn = 3;   // zero‑based index (column D in Excel)

                // Set custom width for the target column and height for the target row
                // (values are in Excel's column width units and points respectively)
                worksheet.Cells.SetColumnWidth(targetColumn, 20);   // wide column
                worksheet.Cells.SetRowHeight(targetRow, 30);       // tall row

                // Convert column width and row height to pixels using worksheet helper methods
                int columnPixelOffset = worksheet.Cells.GetColumnWidthPixel(targetColumn);
                int rowPixelOffset = worksheet.Cells.GetRowHeightPixel(targetRow);

                // Add a rectangle shape (initially placed at cell A1 with zero offsets)
                Shape shape = worksheet.Shapes.AddRectangle(0, 0, 0, 100, 50, 0);

                // Move shape to the target cell range (single cell)
                shape.MoveToRange(targetRow, targetColumn, targetRow, targetColumn);

                // Apply pixel offsets so the shape is positioned absolutely within the cell
                shape.Left = columnPixelOffset / 2;   // example: half the column width from the left edge
                shape.Top = rowPixelOffset / 2;       // example: half the row height from the top edge

                // Validate that the shape is attached to the correct cell
                if (shape.UpperLeftRow == targetRow && shape.UpperLeftColumn == targetColumn)
                {
                    Console.WriteLine($"Shape correctly positioned at row {targetRow}, column {targetColumn}.");
                }
                else
                {
                    Console.WriteLine("Shape positioning mismatch.");
                }

                // Optionally align the shape's top‑right corner to another cell to demonstrate alignment method
                shape.AlignTopRightCorner(targetRow, targetColumn + 2); // align to two columns to the right

                // Save the workbook
                string outputPath = "ShapeAbsolutePositionDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
