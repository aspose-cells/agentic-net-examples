// Title: Get a shape's absolute pixel position and align it to a cell range with Aspose.Cells for .NET (C#)
// Description: This example creates a workbook, adds a rectangle shape, reads its absolute X/Y pixel values (shape.X, shape.Y), converts a target range's Top/Left from points to pixels, and repositions the shape by setting UpperLeftRow, UpperLeftColumn, Top and Left so its top‑left corner matches the range's top‑left corner. The workbook is then saved to verify the alignment.
// Keywords: Aspose.Cells shape position | shape.X shape.Y C# | convert range points to pixels | align shape with cell range | Aspose.Cells rectangle alignment | C# Excel shape coordinates | Aspose.Cells Top Left alignment
// Common Searches: how to get shape pixel coordinates Aspose.Cells .NET | align rectangle shape to specific cells Aspose.Cells | range.Top range.Left conversion to pixels | move shape to cell E5 using Aspose.Cells | C# example shape alignment Excel
// Developer Intent: Retrieve a shape's absolute pixel coordinates and move it so its top‑left corner coincides with the top‑left corner of a given cell range.
// Use Cases: Log the current X and Y pixel location of any worksheet shape. | Place a chart, image, or button precisely over a merged cell block. | Programmatically adjust shape positions after inserting rows or columns. | Create printable reports where shapes must align with table headers.
// AI Prompts: Generate C# code that reads shape.X and shape.Y in pixels and aligns the shape with range "E5:F7" using Aspose.Cells. | Explain the steps to convert Aspose.Cells Range.Top and Range.Left from points to pixels and apply them to a shape's Top and Left properties. | Provide a concise Aspose.Cells example that moves a rectangle shape to the top‑left corner of a target cell range and saves the workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsShapeAlignment
{
    // This example creates a workbook, adds a rectangle shape, reads its absolute X/Y pixel values (shape.X, shape.Y), converts a target range's Top/Left from points to pixels, and repositions the shape by setting UpperLeftRow, UpperLeftColumn, Top and Left so its top‑left corner matches the range's top‑left corner. The workbook is then saved to verify the alignment.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a rectangle shape to the worksheet
                // Parameters: upper left row, upper left column, top offset, left offset, width, height
                Shape shape = worksheet.Shapes.AddRectangle(2, 2, 0, 0, 120, 80);

                // Retrieve the shape's absolute top‑left coordinates (in pixels)
                int shapeAbsoluteX = shape.X; // horizontal offset in pixels
                int shapeAbsoluteY = shape.Y; // vertical offset in pixels
                Console.WriteLine($"Shape absolute position: X = {shapeAbsoluteX} px, Y = {shapeAbsoluteY} px");

                // Define the target cell range we want the shape to align with (e.g., "E5:F7")
                AsposeRange targetRange = worksheet.Cells.CreateRange("E5:F7");

                // Range.Top and Range.Left are returned in points.
                // Convert points to pixels (1 point = 1/72 inch, 1 inch = 96 pixels)
                double pointsToPixels = 96.0 / 72.0; // = 4/3
                int rangeTopPx = (int)Math.Round(targetRange.Top * pointsToPixels);
                int rangeLeftPx = (int)Math.Round(targetRange.Left * pointsToPixels);
                Console.WriteLine($"Target range top‑left: X = {rangeLeftPx} px, Y = {rangeTopPx} px");

                // Align the shape's top‑left corner with the target range's top‑left corner
                shape.UpperLeftRow = targetRange.FirstRow;
                shape.UpperLeftColumn = targetRange.FirstColumn;
                shape.Top = rangeTopPx;   // vertical offset in pixels
                shape.Left = rangeLeftPx; // horizontal offset in pixels

                // Verify the new absolute position
                int newShapeX = shape.X;
                int newShapeY = shape.Y;
                Console.WriteLine($"Shape new absolute position: X = {newShapeX} px, Y = {newShapeY} px");

                // Save the workbook (optional, just to visualize the result)
                workbook.Save("ShapeAlignedWithRange.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
