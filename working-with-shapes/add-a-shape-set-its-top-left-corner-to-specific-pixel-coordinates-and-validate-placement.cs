// Title: C# – Add a Rectangle Shape at Exact Pixel Coordinates and Verify Its Position with Aspose.Cells
// Description: Creates a new workbook, inserts a rectangle shape on the first worksheet at row 5, column 3 with 50 px vertical and 100 px horizontal offsets, checks the Top, Left, UpperLeftRow and UpperLeftColumn properties for accuracy, and saves the file as ShapePlacementDemo.xlsx.
// Keywords: Aspose.Cells C# shape example | AddShape pixel offset | rectangle shape placement Aspose.Cells | Top property Aspose.Cells | Left property Aspose.Cells | UpperLeftRow UpperLeftColumn | worksheet.Shapes.AddShape | verify shape location | GitHub Aspose.Cells sample | code snippet shape positioning
// Common Searches: Aspose.Cells add shape at pixel offset C# | how to set shape top left coordinates Aspose.Cells | validate shape placement properties Aspose.Cells | C# example for worksheet.Shapes.AddShape with pixel values | Aspose.Cells shape positioning verification
// Developer Intent: Insert a shape at precise pixel offsets on a worksheet and confirm that its location properties (Top, Left, UpperLeftRow, UpperLeftColumn) match the intended values.
// Use Cases: Generate reports where logos or diagrams must align to cell boundaries using exact pixel offsets. | Build dashboards that require consistent shape placement relative to specific rows and columns across multiple workbooks. | Create automated tests that add shapes and assert their positional properties to ensure layout logic remains reliable.
// AI Prompts: Write C# code with Aspose.Cells to add an ellipse at row 2, column 4 with a 30 px top offset and 60 px left offset, then verify its placement properties. | Explain the difference between the Top/Left properties and UpperLeftRow/UpperLeftColumn when positioning shapes in Aspose.Cells. | Provide a step‑by‑step guide for confirming shape placement after using worksheet.Shapes.AddShape with pixel offsets.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ShapePlacementDemo
{
    // Creates a new workbook, inserts a rectangle shape on the first worksheet at row 5, column 3 with 50 px vertical and 100 px horizontal offsets, checks the Top, Left, UpperLeftRow and UpperLeftColumn properties for accuracy, and saves the file as ShapePlacementDemo.xlsx.
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a rectangle shape to the worksheet at specific pixel coordinates
                // Parameters: type, topRow, top (pixel offset), leftColumn, left (pixel offset), height, width
                Shape shape = worksheet.Shapes.AddShape(
                    MsoDrawingType.Rectangle, // shape type
                    topRow: 5,                // row index where the shape starts
                    top: 50,                  // vertical offset in pixels from the top of the row
                    leftColumn: 3,            // column index where the shape starts
                    left: 100,                // horizontal offset in pixels from the left of the column
                    height: 200,              // height in pixels
                    width: 150);              // width in pixels

                // Validate placement by checking the shape's properties
                bool isTopCorrect = shape.Top == 50;
                bool isLeftCorrect = shape.Left == 100;
                bool isRowCorrect = shape.UpperLeftRow == 5;
                bool isColumnCorrect = shape.UpperLeftColumn == 3; // Correct property for column index

                Console.WriteLine($"Top correct: {isTopCorrect}");
                Console.WriteLine($"Left correct: {isLeftCorrect}");
                Console.WriteLine($"Row correct: {isRowCorrect}");
                Console.WriteLine($"Column correct: {isColumnCorrect}");

                // Save the workbook
                workbook.Save("ShapePlacementDemo.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
