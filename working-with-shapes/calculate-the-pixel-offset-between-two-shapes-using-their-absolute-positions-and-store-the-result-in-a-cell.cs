// Title: Calculate horizontal and vertical pixel offsets between two worksheet shapes and store them in cells using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that adds two shapes to a worksheet, converts their point coordinates to pixels, computes the X and Y pixel distance between them, and writes the results to cells A1 and B1 with Aspose.Cells. | Show how to retrieve the absolute pixel positions of two shapes, determine the offset values, and save those offsets into specific Excel cells using the Aspose.Cells API.
// Common Searches: Aspose.Cells C# example to find pixel distance between two shapes on a worksheet | how to write shape position offsets to Excel cells using Aspose.Cells .NET | convert shape coordinates from points to pixels in Aspose.Cells C# | determine X and Y offset of two worksheet shapes with Aspose.Cells | store shape offset values in A1 and B1 using Aspose.Cells for .NET
// Tags: calculate shape pixel offset Aspose.Cells C# | write shape offset to worksheet cells Aspose.Cells | convert shape points to pixels Aspose.Cells | add rectangle shapes with pixel positioning Aspose.Cells | retrieve absolute shape positions Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a workbook, adds two rectangle shapes at specific point coordinates, converts those coordinates to pixels, computes the horizontal and vertical pixel offsets between the shapes, writes the X offset to cell A1 and the Y offset to cell B1, and saves the file as ShapeOffset.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Conversion factor: 1 point = 96/72 = 1.33333 pixels
            const double pointsToPixels = 96.0 / 72.0;

            // Shape dimensions (100 points) converted to pixels
            int shapeWidthPixels = (int)(100 * pointsToPixels);
            int shapeHeightPixels = (int)(100 * pointsToPixels);

            // Add the first shape (rectangle) at (50 pt, 80 pt)
            int shape1LeftPixels = (int)(50 * pointsToPixels);
            int shape1TopPixels = (int)(80 * pointsToPixels);
            Shape shape1 = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle,   // shape type
                0, 0,                       // upper-left cell (row, column)
                shape1TopPixels,            // top offset in pixels
                shape1LeftPixels,           // left offset in pixels
                shapeHeightPixels,          // height in pixels
                shapeWidthPixels);          // width in pixels

            // Add the second shape. If the Ellipse type is unavailable, use a rectangle as a fallback.
            int shape2LeftPixels = (int)(200 * pointsToPixels);
            int shape2TopPixels = (int)(150 * pointsToPixels);
            Shape shape2 = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle,   // fallback shape type (replace with MsoDrawingType.Ellipse if supported)
                0, 0,
                shape2TopPixels,
                shape2LeftPixels,
                shapeHeightPixels,
                shapeWidthPixels);

            // Calculate the offset between the two shapes in pixels
            double offsetXPixels = shape2LeftPixels - shape1LeftPixels;
            double offsetYPixels = shape2TopPixels - shape1TopPixels;

            // Store the pixel offsets in cells A1 (X offset) and B1 (Y offset)
            sheet.Cells["A1"].PutValue(offsetXPixels);
            sheet.Cells["B1"].PutValue(offsetYPixels);

            // Save the workbook
            string outputPath = "ShapeOffset.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
