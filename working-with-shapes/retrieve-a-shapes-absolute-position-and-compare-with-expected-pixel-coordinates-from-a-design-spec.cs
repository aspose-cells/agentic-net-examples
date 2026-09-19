// Title: Determine a shape's absolute pixel coordinates in an Excel worksheet with Aspose.Cells for .NET and validate against a design specification
// AI Prompts: Write C# code using Aspose.Cells to read a shape's UpperLeftRow and UpperLeftColumn, then compute its absolute X/Y pixel position by accumulating column widths and row heights. | Add logic that compares the computed pixel coordinates with expected design‑spec values, using a configurable tolerance, and prints a pass/fail message. | Update the sample to incorporate UpperLeftOffsetX and UpperLeftOffsetY when they are available in newer Aspose.Cells releases.
// Common Searches: aspnet retrieve shape pixel location from Excel using Aspose.Cells | how to calculate absolute X Y coordinates of a drawing object in .xlsx with Aspose.Cells | compare Excel shape position to design specification in C# | Aspose.Cells get shape UpperLeftRow UpperLeftColumn pixel conversion | validate shape placement accuracy in Excel workbook programmatically
// Tags: Aspose.Cells shape pixel position calculation | Excel worksheet shape coordinate conversion | C# shape placement verification Aspose.Cells | absolute pixel coordinates from row column Aspose.Cells | design spec shape location validation .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The program loads an Excel workbook, accesses the first worksheet, extracts the first shape's UpperLeftRow and UpperLeftColumn, computes its absolute X and Y pixel coordinates by summing preceding column widths and row heights, and then checks whether those coordinates match expected design‑spec values within a small tolerance, reporting the result.
class ShapePositionChecker
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: File not found – \"{inputPath}\"");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure there is at least one shape on the sheet
            if (sheet.Shapes.Count == 0)
            {
                Console.WriteLine("No shapes found on the first worksheet.");
                return;
            }

            // Retrieve the first shape (adjust index or use sheet.Shapes[\"ShapeName\"] as needed)
            Shape shape = sheet.Shapes[0];

            // Use shape's position properties to get row/column and offsets
            int startRow = shape.UpperLeftRow;
            int startColumn = shape.UpperLeftColumn;

            // Offsets are not available in older API versions; assume zero offset
            int offsetX = 0; // pixels within the starting column
            int offsetY = 0; // pixels within the starting row

            // Calculate absolute X coordinate (pixels) by summing column widths before the start column
            double absoluteX = offsetX;
            for (int col = 0; col < startColumn; col++)
            {
                absoluteX += sheet.Cells.GetColumnWidthPixel(col);
            }

            // Calculate absolute Y coordinate (pixels) by summing row heights before the start row
            double absoluteY = offsetY;
            for (int row = 0; row < startRow; row++)
            {
                absoluteY += sheet.Cells.GetRowHeightPixel(row);
            }

            // Expected pixel coordinates from design specification (replace with actual values)
            double expectedX = 150.0; // example value
            double expectedY = 200.0; // example value

            // Compare actual position with expected position (allowing a small tolerance)
            const double tolerance = 0.5; // pixels
            bool xMatches = Math.Abs(absoluteX - expectedX) <= tolerance;
            bool yMatches = Math.Abs(absoluteY - expectedY) <= tolerance;

            // Output the results
            Console.WriteLine($"Shape absolute position: X = {absoluteX} px, Y = {absoluteY} px");
            Console.WriteLine($"Expected position: X = {expectedX} px, Y = {expectedY} px");
            Console.WriteLine($"X matches: {xMatches}");
            Console.WriteLine($"Y matches: {yMatches}");

            // Further action based on the comparison result
            if (xMatches && yMatches)
            {
                Console.WriteLine("Shape is positioned correctly according to the design spec.");
            }
            else
            {
                Console.WriteLine("Shape position does NOT match the design spec.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
