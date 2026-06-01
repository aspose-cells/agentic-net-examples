using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsShapeAlignment
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a rectangle shape (initial position is arbitrary)
                // Parameters: upper left row, upper left column, top offset, left offset, width, height
                Shape shape = worksheet.Shapes.AddRectangle(5, 2, 0, 0, 100, 50);

                // Define the target cell range to which we want to align the shape
                // Example: range B10:C12 (row index 9, column index 1)
                AsposeRange targetRange = worksheet.Cells.CreateRange("B10:C12");

                // Retrieve the shape's absolute top‑left coordinates (in pixels)
                int shapeAbsoluteX = shape.X; // horizontal offset in pixels
                int shapeAbsoluteY = shape.Y; // vertical offset in pixels

                // Retrieve the target range's top‑left coordinates (in points)
                double rangeTopPoints = targetRange.Top;   // vertical distance in points
                double rangeLeftPoints = targetRange.Left; // horizontal distance in points

                // Output the retrieved coordinates
                Console.WriteLine($"Shape absolute position: X = {shapeAbsoluteX} px, Y = {shapeAbsoluteY} px");
                Console.WriteLine($"Target range top‑left: Left = {rangeLeftPoints} pt, Top = {rangeTopPoints} pt");

                // Align the shape precisely with the target range
                // MoveToRange positions the shape's upper‑left corner to the specified cell indices
                shape.MoveToRange(targetRange.FirstRow, targetRange.FirstColumn,
                                  targetRange.FirstRow, targetRange.FirstColumn);

                // Optionally, adjust pixel offsets within the cell if needed
                shape.X = 0;
                shape.Y = 0;

                // Save the workbook
                string outputPath = "ShapeAlignedToRange.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}