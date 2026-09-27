// Title: How to position a rectangle shape at cell C4 using column width and row height calculations in Aspose.Cells for .NET
// AI Prompts: Compute the total pixel width of columns preceding column C, convert the result to points, and assign it to the Shape.Left property to align the shape with the left edge of the cell. | Sum the pixel heights of rows before row 4, convert the sum to points, set Shape.Top, then specify the shape's Width and Height in points before saving the workbook. | Check that the shape's Left and Top values are within a small tolerance of the calculated offsets, and modify the target row or column indices to reposition the shape as needed.
// Common Searches: Aspose.Cells set shape position based on column index and row index | C# calculate shape left offset using GetColumnWidthPixel Aspose.Cells | Align rectangle shape to specific cell with pixel-to-point conversion in Aspose.Cells | Validate shape alignment tolerance after positioning in Aspose.Cells workbook | Place shape at cell C4 programmatically using Aspose.Cells for .NET
// Tags: shape positioning using column width Aspose.Cells | pixel to point conversion for shape placement .NET | calculate cumulative row height for shape top offset | rectangle shape alignment validation Aspose.Cells | set shape location by cell index Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a workbook, adds a zero‑size rectangle shape, calculates the cumulative pixel widths of columns before column C and pixel heights of rows before row 4, converts these measurements to points, assigns them to the shape's Left and Top properties, sets a fixed width and height, verifies alignment within a small tolerance, and saves the file as ShapePositioned.xlsx.
class ShapePositionExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a rectangle shape with initial size (height, width) set to 0
            // Parameters: type, upperLeftRow, upperLeftColumn, top, left, height, width
            Shape shape = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // shape type
                0,   // upper left row
                0,   // upper left column
                0,   // top offset (points)
                0,   // left offset (points)
                0,   // height (points)
                0);  // width (points)

            // Target cell for positioning: column C (index 2), row 4 (index 3)
            int targetColumn = 2; // zero‑based column index
            int targetRow = 3;    // zero‑based row index

            // Calculate cumulative column width in pixels before the target column
            double offsetX = 0;
            for (int col = 0; col < targetColumn; col++)
            {
                offsetX += sheet.Cells.GetColumnWidthPixel(col);
            }

            // Calculate cumulative row height in pixels before the target row
            double offsetY = 0;
            for (int row = 0; row < targetRow; row++)
            {
                offsetY += sheet.Cells.GetRowHeightPixel(row);
            }

            // Convert pixels to points (1 pixel = 0.75 point at 96 DPI)
            const double pixelsToPoints = 0.75;
            shape.Left = (int)(offsetX * pixelsToPoints);
            shape.Top = (int)(offsetY * pixelsToPoints);

            // Set shape size (e.g., 100x50 points)
            shape.Width = 100;
            shape.Height = 50;

            // Verify alignment within a small tolerance
            const double tolerance = 0.01;
            bool isAligned = Math.Abs(shape.Left - offsetX * pixelsToPoints) < tolerance &&
                             Math.Abs(shape.Top - offsetY * pixelsToPoints) < tolerance;

            Console.WriteLine(isAligned
                ? "Shape is correctly aligned."
                : "Shape alignment mismatch.");

            // Save the workbook
            string outputPath = "ShapePositioned.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
