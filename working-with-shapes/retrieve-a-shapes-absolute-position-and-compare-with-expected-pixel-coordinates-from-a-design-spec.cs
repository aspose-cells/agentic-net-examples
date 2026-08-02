// Title: C# – Verify a Shape’s Absolute Pixel Position and Size in Aspose.Cells with GetActualBox
// Description: Creates a workbook, adds a rectangle shape, sets explicit X/Y offsets, retrieves the shape’s actual pixel box via GetActualBox, compares left, top, width and height against design‑spec values, outputs the results, and saves the file.
// Keywords: Aspose.Cells GetActualBox | shape pixel coordinates C# | Excel shape absolute position | retrieve shape dimensions .NET | compare shape location with design spec | worksheet shape location verification | Aspose.Cells shape size | C# Excel shape positioning | pixel‑perfect shape layout
// Common Searches: Aspose.Cells get shape pixel coordinates | How to read shape position in Excel using C# | Validate shape size with Aspose.Cells | GetActualBox example for shapes | Compare Excel shape location to design mockup
// Developer Intent: Obtain a shape’s exact pixel location and dimensions from a worksheet and confirm they match predefined design specifications.
// Use Cases: Automated QA of generated reports to ensure logos and graphics align with branding guidelines. | Testing Excel templates to verify that inserted charts, images, or callouts occupy the correct grid cells. | Building a layout‑validation tool that flags shapes whose position or size deviates from a UI mockup.
// AI Prompts: Write a C# method that receives expected X, Y, width, and height values, calls GetActualBox on a given shape, and returns true only if all values are within a 0.01‑pixel tolerance. | Generate code that iterates over every shape in a worksheet, logs any shape whose actual pixel box differs from a supplied design spec, and produces a summary report. | Show how to handle floating‑point rounding when comparing GetActualBox results to design coordinates, including configurable tolerance thresholds.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ShapePositionVerification
{
    // Creates a workbook, adds a rectangle shape, sets explicit X/Y offsets, retrieves the shape’s actual pixel box via GetActualBox, compares left, top, width and height against design‑spec values, outputs the results, and saves the file.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            // Parameters: upper left row, upper left column, top offset, left offset, height, width
            Shape shape = worksheet.Shapes.AddRectangle(2, 2, 0, 0, 100, 150);

            // Optionally set explicit X and Y offsets (in pixels) for clarity
            shape.X = 250; // horizontal offset from worksheet left border
            shape.Y = 120; // vertical offset from worksheet top border

            // Retrieve the actual position and size of the shape after any transformations
            // GetActualBox returns an array: [x, y, w, h] in pixels
            float[] actualBox = shape.GetActualBox();

            // Expected design specification (pixel coordinates)
            // These values should be defined according to the design spec you are validating against
            float expectedX = 250f;      // expected left position
            float expectedY = 120f;      // expected top position
            float expectedWidth = 150f;  // expected width
            float expectedHeight = 100f; // expected height

            // Compare actual values with expected values
            bool isXMatch = Math.Abs(actualBox[0] - expectedX) < 0.01f;
            bool isYMatch = Math.Abs(actualBox[1] - expectedY) < 0.01f;
            bool isWidthMatch = Math.Abs(actualBox[2] - expectedWidth) < 0.01f;
            bool isHeightMatch = Math.Abs(actualBox[3] - expectedHeight) < 0.01f;

            // Output comparison results
            Console.WriteLine("Actual Box: Left={0}, Top={1}, Width={2}, Height={3}",
                actualBox[0], actualBox[1], actualBox[2], actualBox[3]);

            Console.WriteLine("Expected Box: Left={0}, Top={1}, Width={2}, Height={3}",
                expectedX, expectedY, expectedWidth, expectedHeight);

            Console.WriteLine("Position X match: {0}", isXMatch);
            Console.WriteLine("Position Y match: {0}", isYMatch);
            Console.WriteLine("Width match: {0}", isWidthMatch);
            Console.WriteLine("Height match: {0}", isHeightMatch);

            // Save the workbook (lifecycle rule: save)
            workbook.Save("ShapePositionVerification.xlsx");
        }
    }
}
