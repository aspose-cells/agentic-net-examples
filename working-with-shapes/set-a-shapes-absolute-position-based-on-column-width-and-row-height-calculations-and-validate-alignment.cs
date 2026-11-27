// Title: Place a rectangle shape at cell C3 by calculating absolute position from column widths and row heights using Aspose.Cells for .NET
// Common Searches: how to calculate shape left and top coordinates from column widths in Aspose.Cells C# | Aspose.Cells set shape position based on cell row height and column width | C# example for aligning a shape to a specific Excel cell using pixel to point conversion in Aspose.Cells | determine absolute pixel offset for a shape in an Aspose.Cells worksheet | validate that a shape is placed at the calculated coordinates with Aspose.Cells .NET
// Tags: calculate shape absolute position Aspose.Cells | shape left top from column width pixels | pixel to point conversion Aspose.Cells | align rectangle shape to specific cell | validate shape coordinates Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to position a rectangle shape at cell C3 by summing column widths and row heights, converting the pixel offsets to points, assigning the Left and Top properties, optionally setting size, validating the placement, and saving the workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];

            // Set custom column widths (in characters) and row heights (in points)
            ws.Cells.SetColumnWidth(0, 12); // Column A
            ws.Cells.SetColumnWidth(1, 15); // Column B
            ws.Cells.SetColumnWidth(2, 20); // Column C
            ws.Cells.SetRowHeight(0, 20);   // Row 1
            ws.Cells.SetRowHeight(1, 30);   // Row 2
            ws.Cells.SetRowHeight(2, 40);   // Row 3

            // Add a rectangle shape directly (returns a Shape object)
            Shape shape = ws.Shapes.AddShape(
                MsoDrawingType.Rectangle,
                0, 0,   // upper left row, column
                0, 0,   // top (pixels), left (pixels) – will be adjusted later
                50, 100 // height, width in points (initial size)
            );

            // Target cell for the shape's top‑left corner (e.g., cell C3)
            int targetColumn = 2; // zero‑based index for column C
            int targetRow = 2;    // zero‑based index for row 3

            // Calculate absolute left position (in pixels) from column start
            double left = 0;
            for (int col = 0; col < targetColumn; col++)
            {
                left += ws.Cells.GetColumnWidthPixel(col);
            }

            // Calculate absolute top position (in pixels) from row start
            double top = 0;
            for (int row = 0; row < targetRow; row++)
            {
                top += ws.Cells.GetRowHeightPixel(row);
            }

            // Convert pixels to points (1 pixel = 72/96 points at 96 DPI)
            const double pixelToPoint = 72.0 / 96.0;
            shape.Left = (int)(left * pixelToPoint);
            shape.Top = (int)(top * pixelToPoint);

            // Optionally set shape size (width & height in points)
            shape.Width = 100;
            shape.Height = 50;

            // Validation: ensure the shape aligns exactly with the calculated coordinates
            if (shape.Left != (int)(left * pixelToPoint))
                throw new Exception("Shape left position does not match calculated value.");
            if (shape.Top != (int)(top * pixelToPoint))
                throw new Exception("Shape top position does not match calculated value.");

            // Save the workbook to a file
            wb.Save("ShapePosition.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
