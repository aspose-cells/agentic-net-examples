// Title: Calculate and verify a shape's absolute pixel position in an Excel worksheet using Aspose.Cells for .NET
// Common Searches: how to get the absolute pixel coordinates of a shape in an Excel file with Aspose.Cells C# | Aspose.Cells C# calculate shape top left position in pixels | compare Excel shape location to design specification using Aspose.Cells | determine shape row and column offsets and convert to pixels with Aspose.Cells | validate shape placement tolerance in an Excel worksheet programmatically
// Tags: Aspose.Cells shape pixel position calculation | Excel shape absolute coordinates .NET | convert shape cell indices to pixels Aspose.Cells | shape location validation against design spec | pixel tolerance comparison for Excel shapes

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an Excel workbook with Aspose.Cells, accesses the first worksheet, retrieves the first shape, computes its top‑left corner's absolute X and Y pixel coordinates by converting the shape's row and column indices using the worksheet's row height and column width in pixels, compares the results to expected design‑spec coordinates within a 0.5‑pixel tolerance, and outputs whether the shape position matches the specification.
class ShapePositionChecker
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index as needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure there is at least one shape in the worksheet
            if (sheet.Shapes.Count == 0)
            {
                Console.WriteLine("No shapes were found in the worksheet.");
                return;
            }

            // Retrieve the first shape (or modify to select by name/index as required)
            Shape shape = sheet.Shapes[0];

            // Get the shape's top‑left cell indices (zero‑based)
            int rowIndex = shape.UpperLeftRow;
            int columnIndex = shape.UpperLeftColumn;

            // Offsets inside the cell (in pixels) – not available in this API version, assume 0
            int rowOffsetPx = 0;
            int columnOffsetPx = 0;

            // Convert column width and row height to pixels
            double columnWidthPx = sheet.Cells.GetColumnWidthPixel(columnIndex);
            double rowHeightPx = sheet.Cells.GetRowHeightPixel(rowIndex);

            // Calculate absolute pixel coordinates of the shape's top‑left corner
            double absoluteX = columnIndex * columnWidthPx + columnOffsetPx;
            double absoluteY = rowIndex * rowHeightPx + rowOffsetPx;

            // Expected pixel coordinates from design specification (replace with actual values)
            double expectedX = 150.0;
            double expectedY = 200.0;

            // Compare the actual position with the expected position
            const double tolerance = 0.5; // allowable deviation in pixels
            bool xMatches = Math.Abs(absoluteX - expectedX) <= tolerance;
            bool yMatches = Math.Abs(absoluteY - expectedY) <= tolerance;

            if (xMatches && yMatches)
            {
                Console.WriteLine("Shape position matches the design specification.");
            }
            else
            {
                Console.WriteLine("Shape position does NOT match the design specification.");
                Console.WriteLine($"Actual X: {absoluteX}, Expected X: {expectedX}");
                Console.WriteLine($"Actual Y: {absoluteY}, Expected Y: {expectedY}");
            }

            // Save the workbook if any modifications were made
            // workbook.Save("output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
