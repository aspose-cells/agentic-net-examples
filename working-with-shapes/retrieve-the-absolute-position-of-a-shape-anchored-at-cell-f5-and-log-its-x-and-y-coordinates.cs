using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a rectangle shape anchored at cell F5 (row index 4, column index 5)
        // Parameters: upperLeftRow, upperLeftColumn, width, height, left offset, top offset
        Shape shape = worksheet.Shapes.AddRectangle(4, 5, 100, 50, 0, 0);
        shape.Name = "MyShape";

        // Locate the shape that is anchored at F5
        Shape targetShape = null;
        foreach (Shape s in worksheet.Shapes)
        {
            if (s.UpperLeftRow == 4 && s.UpperLeftColumn == 5)
            {
                targetShape = s;
                break;
            }
        }

        if (targetShape != null)
        {
            // X and Y give the pixel offsets from the worksheet's left and top borders
            Console.WriteLine("Shape X (pixels): " + targetShape.X);
            Console.WriteLine("Shape Y (pixels): " + targetShape.Y);
        }
        else
        {
            Console.WriteLine("No shape found anchored at cell F5.");
        }

        // Save the workbook (optional, demonstrates the save rule)
        workbook.Save("ShapePositionDemo.xlsx");
    }
}