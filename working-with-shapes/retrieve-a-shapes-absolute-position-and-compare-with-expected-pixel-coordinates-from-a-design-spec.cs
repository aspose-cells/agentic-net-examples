// Title: Check a shape’s absolute pixel position and size with Aspose.Cells GetActualBox in C#
// Description: The sample creates a workbook, adds a rectangle shape, sets explicit X and Y pixel offsets, retrieves the shape’s absolute coordinates and dimensions using GetActualBox, compares them to expected pixel values with a small tolerance, prints the verification result, and saves the file.
// Keywords: Aspose.Cells | GetActualBox | shape position | pixel coordinates | C# | Excel | absolute location | shape size validation | worksheet shape offsets
// Common Searches: Aspose.Cells GetActualBox example | How to get shape pixel coordinates in .NET | Compare Excel shape location with design spec | Retrieve absolute position of rectangle shape using Aspose.Cells | Validate shape dimensions in generated workbook
// Developer Intent: Verify that a shape’s actual X, Y, width, and height match predefined pixel specifications.
// Use Cases: Automated QA of branding elements in program‑generated reports. | Unit testing of Excel templates to ensure decorative shapes are placed correctly before distribution. | Runtime validation that dynamically inserted charts or images appear at the intended location and size.
// AI Prompts: Generate C# code that uses Aspose.Cells to read a shape’s X, Y, width, and height via GetActualBox and compare them to target pixel values with a configurable tolerance. | Explain the calculation behind GetActualBox for shapes added with row/column offsets and pixel offsets in Aspose.Cells. | Provide a logging pattern that records any deviation between actual shape coordinates and design specifications during workbook creation.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample creates a workbook, adds a rectangle shape, sets explicit X and Y pixel offsets, retrieves the shape’s absolute coordinates and dimensions using GetActualBox, compares them to expected pixel values with a small tolerance, prints the verification result, and saves the file.
class ShapePositionCheck
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a rectangle shape at a known location
        // Parameters: upper left row, upper left column, width, height, row offset, column offset
        Shape shape = worksheet.Shapes.AddRectangle(2, 2, 150, 100, 0, 0);

        // Set explicit X and Y offsets (pixels from worksheet borders)
        shape.X = 250; // horizontal offset
        shape.Y = 120; // vertical offset

        // Retrieve the absolute position and size using GetActualBox
        // Returns an array: [x, y, width, height]
        float[] actualBox = shape.GetActualBox();

        // Design specification values to compare against
        int expectedX = 250;          // expected left position in pixels
        int expectedY = 120;          // expected top position in pixels
        int expectedWidth = 150;      // expected width in pixels
        int expectedHeight = 100;     // expected height in pixels

        // Compare actual values with expected spec (allowing a tiny tolerance)
        bool matchesSpec = Math.Abs(actualBox[0] - expectedX) < 0.01 &&
                           Math.Abs(actualBox[1] - expectedY) < 0.01 &&
                           Math.Abs(actualBox[2] - expectedWidth) < 0.01 &&
                           Math.Abs(actualBox[3] - expectedHeight) < 0.01;

        // Output the results
        Console.WriteLine($"Actual Box - Left: {actualBox[0]}, Top: {actualBox[1]}, Width: {actualBox[2]}, Height: {actualBox[3]}");
        Console.WriteLine($"Matches design specification: {matchesSpec}");

        // Save the workbook (optional, demonstrates lifecycle usage)
        workbook.Save("ShapePositionCheck.xlsx");
    }
}
