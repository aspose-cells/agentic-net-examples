using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsZOrderDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            Shape shape = worksheet.Shapes.AddRectangle(10, 10, 100, 100, 0, 0);

            // Set the Z-order position to a specific positive value (e.g., 5)
            shape.ZOrderPosition = 5;

            // Retrieve the Z-order position to confirm it was set correctly
            int currentZOrder = shape.ZOrderPosition;
            Console.WriteLine("Shape ZOrderPosition: " + currentZOrder);

            // Save the workbook (optional, just to complete the lifecycle)
            workbook.Save("ShapeZOrderDemo.xlsx");
        }
    }
}