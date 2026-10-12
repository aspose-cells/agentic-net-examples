// Title: Calculate and verify a shape's absolute pixel position using column widths and row heights in Aspose.Cells for .NET
// AI Prompts: Add a rectangle shape to a worksheet at a specific cell, apply X and Y pixel offsets, then compute its absolute screen coordinates by summing column widths and row heights. | Retrieve the shape's UpperLeftRow and UpperLeftColumn properties, recalculate its absolute pixel location, and compare it with the expected coordinates to confirm correct alignment. | Modify the intra‑cell offset values, re‑run the position calculation, and validate that the shape stays aligned with the target cell.
// Common Searches: Aspose.Cells how to place a shape at a cell with pixel offset | calculate shape absolute position from column width in C# Aspose.Cells | verify that a shape aligns with a specific row and column in Aspose.Cells | get column width in pixels using Aspose.Cells API | determine row height in pixels for shape positioning Aspose.Cells
// Tags: shape absolute pixel positioning Aspose.Cells | column width pixel calculation Aspose.Cells .NET | row height pixel offset Aspose.Cells | shape alignment verification worksheet | add rectangle shape with cell offsets Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates adding a rectangle shape to a worksheet at a given cell with X/Y pixel offsets, computes the expected absolute pixel coordinates by accumulating column widths and row heights, retrieves the shape's stored row/column indices, recalculates its absolute position, compares both values to ensure alignment, and saves the workbook.
class ShapePositionExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Define target cell for shape placement (zero‑based indices)
            int targetRow = 5;      // row 6 in Excel
            int targetColumn = 3;   // column D in Excel

            // Offsets inside the target cell (in pixels)
            int offsetX = 10; // from left edge of the cell
            int offsetY = 15; // from top edge of the cell

            // Add a rectangle shape to the worksheet
            // Parameters: drawing type, upper left row, upper left column,
            // row offset, column offset, width, height
            Shape shape = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle,
                targetRow,
                targetColumn,
                offsetY,   // row offset (vertical)
                offsetX,   // column offset (horizontal)
                100,
                50);

            // -------------------------------------------------
            // Calculate expected absolute position (pixels)
            // -------------------------------------------------
            double absoluteX = 0;
            for (int col = 0; col < targetColumn; col++)
                absoluteX += sheet.Cells.GetColumnWidthPixel(col);
            absoluteX += offsetX; // add intra‑cell offset

            double absoluteY = 0;
            for (int row = 0; row < targetRow; row++)
                absoluteY += sheet.Cells.GetRowHeightPixel(row);
            absoluteY += offsetY; // add intra‑cell offset

            // -------------------------------------------------
            // Retrieve shape's stored position data
            // -------------------------------------------------
            int shapeRow = shape.UpperLeftRow;
            int shapeColumn = shape.UpperLeftColumn;

            // Re‑calculate absolute position from shape's stored data
            double shapeAbsX = 0;
            for (int col = 0; col < shapeColumn; col++)
                shapeAbsX += sheet.Cells.GetColumnWidthPixel(col);
            shapeAbsX += offsetX; // use the same intra‑cell offset used when adding

            double shapeAbsY = 0;
            for (int row = 0; row < shapeRow; row++)
                shapeAbsY += sheet.Cells.GetRowHeightPixel(row);
            shapeAbsY += offsetY; // use the same intra‑cell offset used when adding

            // Compare calculated positions
            bool isXAligned = Math.Abs(absoluteX - shapeAbsX) < 0.01;
            bool isYAligned = Math.Abs(absoluteY - shapeAbsY) < 0.01;

            Console.WriteLine($"Expected absolute X: {absoluteX}  |  Shape absolute X: {shapeAbsX}  |  X aligned: {isXAligned}");
            Console.WriteLine($"Expected absolute Y: {absoluteY}  |  Shape absolute Y: {shapeAbsY}  |  Y aligned: {isYAligned}");

            // Save the workbook (optional)
            string outputPath = "ShapePositionDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
