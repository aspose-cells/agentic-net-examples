// Title: How to get a shape’s absolute row, column and point coordinates with Aspose.Cells for .NET
// Common Searches: Aspose.Cells get shape top and left position in points C# | retrieve shape anchor row and column using Aspose.Cells .NET | convert shape pixel offset to points with Aspose.Cells API | C# code to read shape coordinates from an Excel worksheet using Aspose.Cells | determine absolute position of a shape in Excel file via Aspose.Cells
// Tags: shape absolute position points Aspose.Cells | read shape anchor row column .NET | pixel to point conversion Aspose.Cells | Aspose.Cells shape coordinate extraction | C# Excel shape measurement

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Loads a workbook, accesses the first worksheet, verifies that shapes exist, reads the first shape’s UpperLeftRow, UpperLeftColumn, Top and Left properties (expressed in points), outputs these values to the console, and saves the workbook.
class ShapePositionExample
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one shape
            if (sheet.Shapes.Count > 0)
            {
                // Get the first shape
                Shape shape = sheet.Shapes[0];

                // Row and column where the shape is anchored
                int upperLeftRow = shape.UpperLeftRow;
                int upperLeftColumn = shape.UpperLeftColumn;

                // Offsets are stored in points (1 point = 1/72 inch)
                double shapeTopPoints = shape.Top;   // distance from top of the sheet in points
                double shapeLeftPoints = shape.Left; // distance from left of the sheet in points

                // Output the retrieved values
                Console.WriteLine($"Shape anchored at row {upperLeftRow}, column {upperLeftColumn}");
                Console.WriteLine($"Shape.Top (points): {shapeTopPoints:F2} pt");
                Console.WriteLine($"Shape.Left (points): {shapeLeftPoints:F2} pt");
            }
            else
            {
                Console.WriteLine("No shapes found on the worksheet.");
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
