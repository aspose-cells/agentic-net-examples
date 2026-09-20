// Title: Calculate pixel distance between two worksheet shapes using absolute positions in Aspose.Cells for .NET
// AI Prompts: Write a C# method that receives a Worksheet and two shape indexes and returns the Euclidean distance in pixels between the shapes' centers, using column widths and row heights to compute absolute positions. | Enhance the shape‑distance helper to accept optional offsetX and offsetY values for each shape and incorporate these offsets when calculating absolute pixel coordinates for collision detection. | Develop a C# utility that iterates over all shapes in a worksheet, computes pairwise pixel distances, and reports true when any distance falls below a given collision‑threshold.
// Common Searches: how to get absolute pixel coordinates of a shape in Aspose.Cells C# | compute distance between two Excel shapes using Aspose.Cells API | shape collision detection in Aspose.Cells worksheet .NET | convert shape dimensions from points to pixels in Aspose.Cells | determine center point of a shape in Aspose.Cells for distance calculation
// Tags: calculate shape pixel distance Aspose.Cells | absolute shape position worksheet Aspose.Cells | shape center pixel distance C# | points to pixels conversion Aspose.Cells | shape overlap check Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Provides a C# helper that sums column widths and row heights to obtain absolute top‑left pixel coordinates of two shapes, converts their width/height from points to pixels, finds each center point, and applies the Euclidean formula to return the distance between the shape centers.
public class ShapeCollisionHelper
{
    // Conversion factor from points (used by Aspose.Cells for shape dimensions) to pixels.
    private const double PointsToPixels = 96.0 / 72.0;

    /// <param name="worksheet">Worksheet that contains the shapes.</param>
    /// <param name="shapeIndex1">Zero‑based index of the first shape in the worksheet's Shapes collection.</param>
    /// <param name="shapeIndex2">Zero‑based index of the second shape in the worksheet's Shapes collection.</param>
    /// <returns>Distance in pixels between the two shape centers.</returns>
    public static double GetDistanceBetweenShapes(Worksheet worksheet, int shapeIndex1, int shapeIndex2)
    {
        // Retrieve the shapes from the worksheet.
        Shape shape1 = worksheet.Shapes[shapeIndex1];
        Shape shape2 = worksheet.Shapes[shapeIndex2];

        // Compute absolute top‑left pixel coordinates for each shape.
        double x1 = GetAbsoluteX(worksheet, shape1);
        double y1 = GetAbsoluteY(worksheet, shape1);
        double x2 = GetAbsoluteX(worksheet, shape2);
        double y2 = GetAbsoluteY(worksheet, shape2);

        // Determine the center point of each shape (convert width/height from points to pixels).
        double centerX1 = x1 + shape1.Width * PointsToPixels / 2.0;
        double centerY1 = y1 + shape1.Height * PointsToPixels / 2.0;
        double centerX2 = x2 + shape2.Width * PointsToPixels / 2.0;
        double centerY2 = y2 + shape2.Height * PointsToPixels / 2.0;

        // Euclidean distance between the two centers.
        double deltaX = centerX2 - centerX1;
        double deltaY = centerY2 - centerY1;
        return Math.Sqrt(deltaX * deltaX + deltaY * deltaY);
    }

    // Helper: calculates the absolute X (pixel) position of a shape's top‑left corner.
    private static double GetAbsoluteX(Worksheet ws, Shape shape)
    {
        double x = 0.0;

        // Sum widths of all columns that are completely to the left of the shape.
        for (int col = 0; col < shape.UpperLeftColumn; col++)
        {
            x += ws.Cells.GetColumnWidthPixel(col);
        }

        // Offsets inside the starting cell are not available in older API versions;
        // they are treated as zero for compatibility.
        return x;
    }

    // Helper: calculates the absolute Y (pixel) position of a shape's top‑left corner.
    private static double GetAbsoluteY(Worksheet ws, Shape shape)
    {
        double y = 0.0;

        // Sum heights of all rows that are completely above the shape.
        for (int row = 0; row < shape.UpperLeftRow; row++)
        {
            y += ws.Cells.GetRowHeightPixel(row);
        }

        // Offsets inside the starting cell are not available in older API versions;
        // they are treated as zero for compatibility.
        return y;
    }
}

public class Program
{
    public static void Main()
    {
        try
        {
            string filePath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException.
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return;
            }

            // Load the workbook.
            Workbook wb = new Workbook(filePath);
            Worksheet ws = wb.Worksheets[0];

            // Ensure there are at least two shapes to compare.
            if (ws.Shapes.Count < 2)
            {
                Console.WriteLine("The worksheet does not contain at least two shapes.");
                return;
            }

            // Calculate distance between the first two shapes.
            double distance = ShapeCollisionHelper.GetDistanceBetweenShapes(ws, 0, 1);
            Console.WriteLine($"Distance between shape 0 and shape 1: {distance} pixels");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors.
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
