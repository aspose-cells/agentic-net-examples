// Title: Compute Euclidean distance between two worksheet shapes using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# get distance between two shapes on a worksheet | calculate center point of an Excel shape with Aspose.Cells | how to detect collision between shapes in an Excel file using Aspose.Cells | C# example for measuring separation of shape objects in an Aspose.Cells workbook | retrieve shape left top width height properties Aspose.Cells to compute spacing
// Tags: shape center distance Aspose.Cells | worksheet shape absolute position C# | collision detection Excel shapes Aspose.Cells | euclidean separation of shapes Aspose.Cells | retrieve shape dimensions Aspose.Cells .NET

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ShapeDistanceDemo
{
    // The sample loads an Excel workbook, accesses the first worksheet's shape collection, obtains two shapes (by name or the first two entries), computes each shape's center using its Left, Top, Width, and Height properties, and returns the Euclidean distance between the centers for collision detection purposes.
    class Program
    {
        /// <param name="shapeA">First shape.</param>
        /// <param name="shapeB">Second shape.</param>
        /// <returns>Distance in points.</returns>
        static double CalculateShapeCenterDistance(Shape shapeA, Shape shapeB)
        {
            // Center of shape A
            double centerAX = shapeA.Left + shapeA.Width / 2.0;
            double centerAY = shapeA.Top + shapeA.Height / 2.0;

            // Center of shape B
            double centerBX = shapeB.Left + shapeB.Width / 2.0;
            double centerBY = shapeB.Top + shapeB.Height / 2.0;

            // Euclidean distance
            double deltaX = centerAX - centerBX;
            double deltaY = centerAY - centerBY;
            return Math.Sqrt(deltaX * deltaX + deltaY * deltaY);
        }

        static void Main(string[] args)
        {
            // Path to the Excel file that contains the shapes.
            // Ensure the file exists at this location before running the program.
            string workbookPath = "input.xlsx";

            if (!System.IO.File.Exists(workbookPath))
            {
                Console.WriteLine($"Error: Workbook file not found at '{workbookPath}'.");
                return;
            }

            try
            {
                // Load the workbook.
                Workbook workbook = new Workbook(workbookPath);

                // Use the first worksheet for this demo.
                Worksheet worksheet = workbook.Worksheets[0];

                // Retrieve the shapes collection.
                ShapeCollection shapes = worksheet.Shapes;

                if (shapes.Count < 2)
                {
                    Console.WriteLine("Error: The worksheet must contain at least two shapes for distance calculation.");
                    return;
                }

                // Attempt to locate shapes by name; fallback to the first two shapes if not found.
                Shape shape1 = null;
                Shape shape2 = null;

                // Replace these names with the actual names of your shapes if known.
                string shapeNameA = "Shape 1";
                string shapeNameB = "Shape 2";

                try { shape1 = shapes[shapeNameA]; } catch { /* ignore */ }
                try { shape2 = shapes[shapeNameB]; } catch { /* ignore */ }

                if (shape1 == null || shape2 == null)
                {
                    Console.WriteLine("Info: Named shapes not found. Using the first two shapes in the collection.");
                    shape1 = shapes[0];
                    shape2 = shapes[1];
                }

                // Output basic shape information.
                Console.WriteLine($"Shape A: Name='{shape1.Name}', Top={shape1.Top}, Left={shape1.Left}, Width={shape1.Width}, Height={shape1.Height}");
                Console.WriteLine($"Shape B: Name='{shape2.Name}', Top={shape2.Top}, Left={shape2.Left}, Width={shape2.Width}, Height={shape2.Height}");

                // Calculate and display the distance.
                double distance = CalculateShapeCenterDistance(shape1, shape2);
                Console.WriteLine($"Distance between shape centers: {distance:F2} points");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An exception occurred: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
