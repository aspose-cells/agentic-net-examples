// Title: Determine a named shape's absolute pixel coordinates in an Excel worksheet using Aspose.Cells for .NET and validate against design spec
// AI Prompts: Write C# code with Aspose.Cells to locate a shape by its name, calculate its absolute X and Y pixel positions by summing preceding row heights and column widths, and store the results. | Add logic that compares the calculated pixel coordinates with expected design values using a 1‑pixel tolerance and prints a match or mismatch message.
// Common Searches: aspocells get absolute pixel position of a shape in C# | how to calculate Excel shape coordinates in pixels using Aspose.Cells | verify Excel shape location matches design spec with .NET | C# sum row heights and column widths to find shape top‑left pixel offset | Aspose.Cells shape position tolerance check example
// Tags: retrieve shape absolute pixel coordinates Aspose.Cells | calculate shape top-left offset from rows columns .NET | compare shape position with design spec using tolerance | excel worksheet shape location verification C# | sum row heights column widths pixel calculation Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The program loads an Excel workbook, finds a shape by its name, computes the shape's absolute X and Y pixel coordinates by aggregating the heights of preceding rows and the widths of preceding columns, then checks those coordinates against expected design values with a 1‑pixel tolerance, reporting a match or mismatch.
class ShapePositionChecker
{
    static void Main()
    {
        const string inputFile = "input.xlsx";
        const string shapeName = "MyShape"; // replace with actual shape name

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputFile))
        {
            Console.WriteLine($"Error: The file \"{inputFile}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputFile);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Retrieve the shape by name
            Shape shape = worksheet.Shapes[shapeName];
            if (shape == null)
            {
                Console.WriteLine($"Error: Shape \"{shapeName}\" not found in the worksheet.");
                return;
            }

            // Get the cell indices where the shape's top‑left corner is anchored
            int rowIndex = shape.UpperLeftRow;
            int columnIndex = shape.UpperLeftColumn;

            // Offsets inside the cell (in pixels) – not available in older API versions, default to 0
            int rowOffsetPixels = 0;
            int columnOffsetPixels = 0;

            // Calculate absolute Y coordinate (pixels) by summing heights of all preceding rows
            double absoluteY = 0;
            for (int i = 0; i < rowIndex; i++)
            {
                absoluteY += worksheet.Cells.GetRowHeightPixel(i);
            }
            absoluteY += rowOffsetPixels;

            // Calculate absolute X coordinate (pixels) by summing widths of all preceding columns
            double absoluteX = 0;
            for (int i = 0; i < columnIndex; i++)
            {
                absoluteX += worksheet.Cells.GetColumnWidthPixel(i);
            }
            absoluteX += columnOffsetPixels;

            // Expected pixel coordinates from the design specification
            int expectedX = 150; // replace with actual expected X
            int expectedY = 200; // replace with actual expected Y

            // Compare with a tolerance of 1 pixel
            const double tolerance = 1.0;
            bool xMatches = Math.Abs(absoluteX - expectedX) <= tolerance;
            bool yMatches = Math.Abs(absoluteY - expectedY) <= tolerance;

            // Output the result
            if (xMatches && yMatches)
            {
                Console.WriteLine("Shape position matches the expected coordinates.");
            }
            else
            {
                Console.WriteLine($"Shape position mismatch. Actual: ({absoluteX:F2}, {absoluteY:F2}) " +
                                  $"Expected: ({expectedX}, {expectedY})");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
