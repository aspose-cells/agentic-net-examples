using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsZOrderDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            // Parameters: upper left row, upper left column, top offset, left offset, height, width
            Shape shape = worksheet.Shapes.AddRectangle(5, 5, 0, 0, 100, 200);

            // Set a large positive Z-order position to ensure the shape is on top of all other objects
            shape.ZOrderPosition = 10000; // large integer

            // Optional: verify the Z-order position by writing it to the console
            Console.WriteLine("Shape ZOrderPosition set to: " + shape.ZOrderPosition);

            // Save the workbook to a file
            workbook.Save("ShapeZOrderTopDemo.xlsx");
        }
    }
}