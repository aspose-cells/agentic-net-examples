// Title: Compute pixel offset between two worksheet shapes and write the result to a cell – Aspose.Cells for .NET
// Description: This C# example creates a workbook, adds two rectangle shapes, sets their absolute Left and Top pixel coordinates, calculates the horizontal and vertical differences, derives the Euclidean pixel offset, and stores the distance (plus individual deltas) in cells A1‑A3 before saving the file.
// Keywords: Aspose.Cells shape offset | pixel distance between shapes | shape Left Top properties | C# calculate shape distance | store calculation in Excel cell | Euclidean offset Aspose.Cells
// Common Searches: Aspose.Cells get pixel offset of shapes | calculate distance between two shapes in Excel using .NET | write shape offset result to a worksheet cell | how to use Left and Top properties of Aspose.Cells shapes | measure layout spacing with Aspose.Cells
// Developer Intent: Find the pixel offset between two worksheet shapes and record the value in a cell.
// Use Cases: Validate spacing between diagram elements by measuring exact pixel gaps. | Generate a report that lists horizontal, vertical, and total pixel distances for annotated objects. | Automate alignment checks, flagging shapes that exceed a predefined offset threshold.
// AI Prompts: Generate C# code with Aspose.Cells that computes horizontal and vertical pixel differences between two shapes and places the Euclidean distance in cell B5. | Show how to iterate over all shapes in a worksheet, identify the pair with the maximum pixel offset, and write the distance to a summary sheet. | Explain the conversion from shape column/row offsets to absolute pixel coordinates before calculating distance using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapeOffsetDemo
{
    // This C# example creates a workbook, adds two rectangle shapes, sets their absolute Left and Top pixel coordinates, calculates the horizontal and vertical differences, derives the Euclidean pixel offset, and stores the distance (plus individual deltas) in cells A1‑A3 before saving the file.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add first rectangle shape at (column 2, row 3) with size 100x50 pixels
            // Parameters: upper left column, upper left row, upper left offset X, upper left offset Y, width, height
            Shape shape1 = worksheet.Shapes.AddRectangle(2, 3, 0, 0, 100, 50);
            // Set explicit pixel positions (optional, demonstrates usage of Left/Top)
            shape1.Left = 150;   // horizontal offset from left column in pixels
            shape1.Top = 200;    // vertical offset from top row in pixels

            // Add second rectangle shape at (column 5, row 6) with size 80x40 pixels
            Shape shape2 = worksheet.Shapes.AddRectangle(5, 6, 0, 0, 80, 40);
            shape2.Left = 300;
            shape2.Top = 350;

            // Calculate horizontal and vertical pixel differences
            int deltaX = shape2.Left - shape1.Left;   // positive if shape2 is to the right of shape1
            int deltaY = shape2.Top - shape1.Top;    // positive if shape2 is below shape1

            // Calculate Euclidean distance (pixel offset) between the two shapes
            double distance = Math.Sqrt(deltaX * deltaX + deltaY * deltaY);

            // Store the calculated distance in cell A1
            worksheet.Cells["A1"].PutValue(distance);

            // Optionally, store individual deltas for reference
            worksheet.Cells["A2"].PutValue(deltaX); // horizontal offset
            worksheet.Cells["A3"].PutValue(deltaY); // vertical offset

            // Save the workbook
            workbook.Save("ShapeOffsetResult.xlsx");
        }
    }
}
