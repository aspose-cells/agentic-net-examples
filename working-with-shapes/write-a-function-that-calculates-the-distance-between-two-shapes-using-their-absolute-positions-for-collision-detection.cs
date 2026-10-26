// Title: Calculate pixel distance between two worksheet shapes for collision detection with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# get distance between two shapes on a worksheet | how to compute Euclidean distance of shape centers in Excel using Aspose.Cells | convert shape position from points to pixels in Aspose.Cells .NET | detect shape collision in Excel workbook with Aspose.Cells API | retrieve shape by name and measure distance in Aspose.Cells C# example | calculate distance between rectangle shapes in an Excel file using Aspose.Cells
// Tags: shape center calculation Aspose.Cells | excel points to pixel conversion C# | worksheet shape distance measurement | collision detection for Excel shapes Aspose | retrieve shapes by name Aspose.Cells API | euclidean distance between shape centers C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The GetShapesDistance method fetches two shapes from a worksheet by their names, converts their Left, Top, Width, and Height from Excel points to screen pixels (assuming 96 DPI), determines each shape's center coordinates, and returns the Euclidean distance between the centers in pixels. It throws an ArgumentException if either shape cannot be found.
public static class ShapeCollisionHelper
{
    /// <param name="worksheet">The worksheet containing the shapes.</param>
    /// <param name="shapeName1">The name (or index) of the first shape.</param>
    /// <param name="shapeName2">The name (or index) of the second shape.</param>
    /// <returns>The distance between the two shape centers in pixels.</returns>
    public static double GetShapesDistance(Worksheet worksheet, string shapeName1, string shapeName2)
    {
        // Retrieve the shapes from the worksheet.
        Shape shape1 = worksheet.Shapes[shapeName1];
        Shape shape2 = worksheet.Shapes[shapeName2];

        if (shape1 == null || shape2 == null)
            throw new ArgumentException("One or both shapes could not be found on the worksheet.");

        // Convert Excel points (1/72 inch) to screen pixels (assuming 96 DPI).
        double PointsToPixels(double points) => points * 96.0 / 72.0;

        // Compute the absolute center point of a shape in pixel coordinates.
        (double X, double Y) GetShapeCenter(Shape shape)
        {
            double leftPx   = PointsToPixels(shape.Left);
            double topPx    = PointsToPixels(shape.Top);
            double widthPx  = PointsToPixels(shape.Width);
            double heightPx = PointsToPixels(shape.Height);

            double centerX = leftPx + widthPx / 2.0;
            double centerY = topPx + heightPx / 2.0;

            return (centerX, centerY);
        }

        var center1 = GetShapeCenter(shape1);
        var center2 = GetShapeCenter(shape2);

        // Euclidean distance between the two centers.
        double deltaX = center2.X - center1.X;
        double deltaY = center2.Y - center1.Y;
        return Math.Sqrt(deltaX * deltaX + deltaY * deltaY);
    }
}

public class Program
{
    public static void Main()
    {
        try
        {
            string filePath = "sample.xlsx";

            // Load existing workbook if present; otherwise create a new one.
            Workbook workbook = File.Exists(filePath) ? new Workbook(filePath) : new Workbook();

            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure at least two shapes exist for the demo.
            if (worksheet.Shapes.Count < 2)
            {
                Shape s1 = worksheet.Shapes.AddShape(MsoDrawingType.Rectangle, 1, 1, 0, 0, 100, 50);
                s1.Name = "Shape1";

                Shape s2 = worksheet.Shapes.AddShape(MsoDrawingType.Rectangle, 5, 5, 0, 0, 80, 40);
                s2.Name = "Shape2";
            }

            double distance = ShapeCollisionHelper.GetShapesDistance(worksheet, "Shape1", "Shape2");
            Console.WriteLine($"Distance between shapes: {distance} pixels");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
