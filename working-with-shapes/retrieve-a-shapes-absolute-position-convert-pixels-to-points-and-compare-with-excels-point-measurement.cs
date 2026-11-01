// Title: Get a worksheet shape's absolute top and left coordinates in points and convert them to pixels using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that loads a workbook, accesses the first shape on the first worksheet, and prints its Top and Left values in both points and pixels. | Create a helper method in C# that returns a shape's position as a struct containing the top/left measurements in points and the equivalent pixel values using Aspose.Cells. | Extend the sample to also output the shape's width and height, converting those dimensions from points to pixels with the 96/72 conversion factor.
// Common Searches: Aspose.Cells get shape top coordinate in points C# | convert shape left position from points to pixels Aspose.Cells | how to read absolute position of Excel drawing shape using Aspose.Cells .NET | pixel conversion formula for Excel shape measurements C# Aspose
// Tags: Aspose.Cells shape position points to pixels conversion | C# retrieve worksheet shape coordinates | Aspose.Cells get shape top left values | Excel shape measurement conversion .NET | Aspose.Cells shape dimensions conversion points pixels

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads a workbook, accesses the first shape on the first worksheet, reads its Top and Left properties (stored in points), converts those values to pixels using the 96/72 ratio, prints both units, and saves the workbook.
class ShapePositionExample
{
    static void Main()
    {
        try
        {
            // Load an existing workbook if present; otherwise create a new one.
            const string inputPath = "input.xlsx";
            Workbook workbook = File.Exists(inputPath) ? new Workbook(inputPath) : new Workbook();

            // Access the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure there is at least one shape on the sheet.
            if (sheet.Shapes.Count == 0)
            {
                Console.WriteLine("No shapes found on the first worksheet.");
                return;
            }

            // Retrieve the first shape.
            Shape shape = sheet.Shapes[0];

            // Excel stores the shape's top/left position in points.
            double shapeTopPoints = shape.Top;
            double shapeLeftPoints = shape.Left;

            // Convert points to pixels (1 point = 1/72 inch, 1 pixel = 1/96 inch).
            double shapeTopPixels = shapeTopPoints * 96.0 / 72.0;
            double shapeLeftPixels = shapeLeftPoints * 96.0 / 72.0;

            Console.WriteLine($"Shape top: {shapeTopPoints:F2} pts ({shapeTopPixels:F0} px)");
            Console.WriteLine($"Shape left: {shapeLeftPoints:F2} pts ({shapeLeftPixels:F0} px)");

            // Save the workbook.
            const string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
