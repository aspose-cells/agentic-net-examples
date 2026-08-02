// Title: Add a rectangle shape, set pixel‑based Top/Left coordinates, and verify placement with Aspose.Cells for .NET (C#)
// Description: Creates a new workbook, inserts a rectangle shape on the first worksheet, assigns exact pixel values to the shape's Top and Left properties, checks that the shape is positioned as expected, outputs the result, and saves the file as ShapePlacement.xlsx.
// Keywords: Aspose.Cells shape positioning | C# set shape top left pixels | Aspose.Cells rectangle shape | validate shape coordinates | shape.Top property | shape.Left property | Excel shape pixel offset
// Common Searches: Aspose.Cells set shape position in pixels | C# place rectangle shape at exact location in Excel | how to get shape coordinates Aspose.Cells | verify shape placement programmatically | pixel offset for Excel shapes Aspose.Cells
// Developer Intent: Add a rectangle shape to a worksheet, move it to precise pixel coordinates, and programmatically confirm that the placement matches the requested values.
// Use Cases: Insert a company logo at a fixed offset in generated financial reports. | Align a watermark shape consistently across exported PDF files. | Validate shape locations when building dynamic Excel templates for automated reporting.
// AI Prompts: Generate C# code with Aspose.Cells that adds a circle shape, sets Top to 30 px and Left to 80 px, and throws an exception if the coordinates differ. | Create a reusable C# method that receives pixel X/Y values, positions any given shape using its Top and Left properties, and returns a boolean indicating success. | Explain how to convert worksheet row/column indices to pixel offsets for accurate shape positioning in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ShapePlacementDemo
{
    // Creates a new workbook, inserts a rectangle shape on the first worksheet, assigns exact pixel values to the shape's Top and Left properties, checks that the shape is positioned as expected, outputs the result, and saves the file as ShapePlacement.xlsx.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Access the shape collection of the worksheet
            ShapeCollection shapes = worksheet.Shapes;

            // Add a rectangle shape with initial zero offsets
            // Parameters: type, topRow, top (pixel), leftColumn, left (pixel), height (pixel), width (pixel)
            Shape shape = shapes.AddShape(
                MsoDrawingType.Rectangle, // shape type
                0,    // top row index
                0,    // vertical offset in pixels (will be updated)
                0,    // left column index
                0,    // horizontal offset in pixels (will be updated)
                100,  // height in pixels
                200); // width in pixels

            // Desired top-left corner coordinates (in pixels)
            int desiredTop = 50;   // vertical offset from the top of the worksheet
            int desiredLeft = 120; // horizontal offset from the left of the worksheet

            // Set the shape's position
            shape.Top = desiredTop;
            shape.Left = desiredLeft;

            // Validate that the shape is placed at the expected coordinates
            if (shape.Top == desiredTop && shape.Left == desiredLeft)
            {
                Console.WriteLine("Shape placed correctly at Top={0}, Left={1}.", shape.Top, shape.Left);
            }
            else
            {
                Console.WriteLine("Shape placement mismatch. Current Top={0}, Left={1}.", shape.Top, shape.Left);
            }

            // Save the workbook
            workbook.Save("ShapePlacement.xlsx");
        }
    }
}
