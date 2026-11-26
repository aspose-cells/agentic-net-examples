// Title: Calculate pixel offset between two worksheet shapes and store the distance in a cell with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells calculate distance between two shapes in pixels | How to get pixel offset of shapes and write result to Excel cell using Aspose.Cells | Aspose.Cells shape position difference row column to pixel conversion example | Store computed shape offset value in cell A1 with Aspose.Cells C# | Determine absolute position offset between rectangle and ellipse shapes in Aspose.Cells
// Tags: shape pixel offset calculation Aspose.Cells | convert shape row column difference to pixels C# | write computed distance to worksheet cell Aspose.Cells | rectangle ellipse position comparison Aspose.Cells | absolute shape position handling Aspose.Cells .NET

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a workbook, adds a rectangle and an ellipse shape with specific row, column, and pixel offsets, computes the absolute row and column differences, converts them to an approximate pixel distance, writes the distance into cell A1, and saves the workbook as ShapeOffsetResult.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // -------------------------------------------------
            // Add first shape (a rectangle) and set its position
            // -------------------------------------------------
            Shape rectShape = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // shape type
                2,    // upper left row
                1,    // upper left column
                5,    // row offset in pixels
                10,   // column offset in pixels
                30,   // height in points
                60);  // width in points

            // -------------------------------------------------
            // Add second shape (an ellipse) and set its position
            // -------------------------------------------------
            // If the current Aspose.Cells version does not expose MsoDrawingType.Ellipse,
            // use its underlying integer value (3) which corresponds to an ellipse.
            Shape ellipseShape = sheet.Shapes.AddShape(
                (MsoDrawingType)3, // ellipse shape type
                5,    // upper left row
                4,    // upper left column
                12,   // row offset in pixels
                8,    // column offset in pixels
                40,   // height in points
                80);  // width in points

            // -------------------------------------------------
            // Calculate distance between the two shapes (cell based)
            // -------------------------------------------------
            // Use cell indices because pixel‑offset properties may not be available
            int deltaRows = Math.Abs(rectShape.UpperLeftRow - ellipseShape.UpperLeftRow);
            int deltaCols = Math.Abs(rectShape.UpperLeftColumn - ellipseShape.UpperLeftColumn);

            // Approximate conversion factors (can be adjusted as needed)
            const double pixelsPerRow = 15.0;
            const double pixelsPerColumn = 64.0;

            double pixelDistance = Math.Sqrt(
                Math.Pow(deltaRows * pixelsPerRow, 2) +
                Math.Pow(deltaCols * pixelsPerColumn, 2));

            // -------------------------------------------------
            // Store the result in cell A1
            // -------------------------------------------------
            sheet.Cells["A1"].PutValue(pixelDistance);

            // -------------------------------------------------
            // Save the workbook
            // -------------------------------------------------
            string outputPath = "ShapeOffsetResult.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'. Distance: {pixelDistance}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
