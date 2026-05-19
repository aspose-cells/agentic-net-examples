using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapeZOrderDemo
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

            // Bring the shape to the front (positive value moves forward)
            shape.ToFrontOrBack(1);

            // Capture the Z-order position after moving to front
            int zOrder = shape.ZOrderPosition;

            // Log the Z-order index to a text file
            string logPath = "ShapeZOrderLog.txt";
            using (StreamWriter writer = new StreamWriter(logPath, append: true))
            {
                writer.WriteLine($"{DateTime.Now}: Shape ZOrderPosition = {zOrder}");
            }

            // Save the workbook (optional, just to demonstrate lifecycle usage)
            workbook.Save("ShapeZOrderDemo.xlsx");
        }
    }
}