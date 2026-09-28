// Title: Retrieve a shape's absolute top and left coordinates in points, convert them to pixel offsets, and compare the values using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that opens an Excel workbook with Aspose.Cells, reads the first shape's Top and Left properties in points, converts those values to pixel offsets (using 1 point = 1/0.75 pixels), converts the pixel values back to points, and prints the original, converted, and difference values. | Show how to calculate pixel offsets from a shape's point coordinates and verify the conversion accuracy with Aspose.Cells in a .NET application.
// Common Searches: how to get shape top left position in points with Aspose.Cells C# | convert shape coordinates from points to pixels using Aspose.Cells | compare Aspose.Cells shape point values with pixel offsets in Excel | C# Aspose.Cells retrieve absolute position of a drawing shape | pixel to point conversion for Excel shapes in .NET
// Tags: Aspose.Cells shape position points | shape point to pixel conversion Aspose.Cells | retrieve shape top left coordinates C# | Excel shape absolute position Aspose.Cells | pixel offset verification Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an Excel workbook, accesses the first worksheet, obtains the first shape, reads its Top and Left values in points, converts those values to pixel offsets using the 1 point = 1/0.75 pixel ratio, converts the pixel offsets back to points, calculates the differences, and writes all results to the console.
class ShapePositionExample
{
    static void Main()
    {
        string filePath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"File not found: {filePath}");
            return;
        }

        try
        {
            // Load the workbook from the specified file
            Workbook workbook = new Workbook(filePath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure there is at least one shape on the worksheet
            if (sheet.Shapes.Count == 0)
            {
                Console.WriteLine("No shapes found on the worksheet.");
                return;
            }

            // Get the first shape
            Shape shape = sheet.Shapes[0];

            // Retrieve the shape's absolute position in points
            double topInPoints = shape.Top;
            double leftInPoints = shape.Left;

            // Convert points to pixel offsets (1 point = 1/0.75 pixels)
            int topOffsetPixels = (int)Math.Round(topInPoints / 0.75);
            int leftOffsetPixels = (int)Math.Round(leftInPoints / 0.75);

            // Convert pixel offsets back to points for comparison
            double topOffsetPoints = topOffsetPixels * 0.75;
            double leftOffsetPoints = leftOffsetPixels * 0.75;

            // Compare the Aspose.Cells point values with the converted pixel values
            double topDifference = Math.Abs(topInPoints - topOffsetPoints);
            double leftDifference = Math.Abs(leftInPoints - leftOffsetPoints);

            // Output the results
            Console.WriteLine($"Shape Top (points): {topInPoints}");
            Console.WriteLine($"Shape Left (points): {leftInPoints}");
            Console.WriteLine($"Top offset (pixels) -> points: {topOffsetPixels} px = {topOffsetPoints} pt");
            Console.WriteLine($"Left offset (pixels) -> points: {leftOffsetPixels} px = {leftOffsetPoints} pt");
            Console.WriteLine($"Difference in Top: {topDifference} points");
            Console.WriteLine($"Difference in Left: {leftDifference} points");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
