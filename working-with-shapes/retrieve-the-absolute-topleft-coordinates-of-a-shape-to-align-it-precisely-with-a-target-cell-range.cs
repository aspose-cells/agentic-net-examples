// Title: Get Shape Absolute Position and Align It to a Cell Range with Aspose.Cells for .NET
// Description: Creates a workbook, adds a rectangle shape, reads its absolute X/Y pixel coordinates, converts a target cell range's Top and Left from points to pixels, and sets the shape's X and Y so it aligns exactly with the range before saving the file.
// Keywords: Aspose.Cells shape position | shape X Y properties | align shape to cell range | convert points to pixels Aspose.Cells | C# Aspose.Cells shape alignment | absolute shape coordinates
// Common Searches: how to get shape pixel position Aspose.Cells | align rectangle shape with specific cells .NET | convert range.Top and range.Left to pixels | set shape.X and shape.Y based on cell coordinates | Aspose.Cells shape alignment example
// Developer Intent: Retrieve a shape’s absolute top‑left pixel coordinates and move it so the shape lines up with the top‑left corner of a specified cell range.
// Use Cases: Place a logo shape precisely at the start of a header row. | Position comment or note shapes next to data tables automatically. | Align multiple chart or image shapes with their corresponding data blocks when generating reports.
// AI Prompts: Write C# code that reads a shape’s X and Y values and aligns it to the top‑left corner of range "B2:C3" using Aspose.Cells. | Explain the conversion from Aspose.Cells range.Top/Left points to pixels and how to apply them to shape positioning. | Provide robust error‑handling patterns for aligning shapes to cell ranges in Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsShapeAlignment
{
    // Creates a workbook, adds a rectangle shape, reads its absolute X/Y pixel coordinates, converts a target cell range's Top and Left from points to pixels, and sets the shape's X and Y so it aligns exactly with the range before saving the file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (lifecycle rule: create)
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a rectangle shape to the worksheet
                // Parameters: upper left row, upper left column, top offset, left offset, width, height
                Shape shape = worksheet.Shapes.AddRectangle(1, 1, 0, 0, 100, 50);

                // Retrieve the shape's absolute top‑left coordinates (pixels from worksheet origin)
                int shapeAbsoluteX = (int)shape.X; // horizontal offset from worksheet left border
                int shapeAbsoluteY = (int)shape.Y; // vertical offset from worksheet top border

                Console.WriteLine($"Shape absolute position: X = {shapeAbsoluteX} px, Y = {shapeAbsoluteY} px");

                // Define the target cell range to which the shape should be aligned
                AsposeRange targetRange = worksheet.Cells.CreateRange("C5:D7");

                // Retrieve the range's top‑left position in points (distance from row 1 / column A)
                double rangeTopPoints = targetRange.Top;   // points from top edge of row 1
                double rangeLeftPoints = targetRange.Left; // points from left edge of column A

                // Convert points to pixels (1 point = 1/72 inch, 1 inch = 96 pixels by default)
                const double pointsToPixels = 96.0 / 72.0; // 1.33333
                int rangeTopPixels = (int)Math.Round(rangeTopPoints * pointsToPixels);
                int rangeLeftPixels = (int)Math.Round(rangeLeftPoints * pointsToPixels);

                Console.WriteLine($"Target range top‑left: X = {rangeLeftPixels} px, Y = {rangeTopPixels} px");

                // Align the shape precisely with the target range using absolute coordinates
                shape.X = rangeLeftPixels;
                shape.Y = rangeTopPixels;

                // Save the workbook (lifecycle rule: save)
                string outputPath = "ShapeAlignedWithRange.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
