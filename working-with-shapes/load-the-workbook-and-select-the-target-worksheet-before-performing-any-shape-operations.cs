using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapeDemo
{
    public class Program
    {
        public static void Main()
        {
            // Load an existing workbook from file (uses Workbook(string) constructor)
            Workbook workbook = new Workbook("input.xlsx");

            // Select the target worksheet (first worksheet in this example)
            Worksheet targetSheet = workbook.Worksheets[0];

            // Perform shape operations on the selected worksheet
            // Add a rectangle shape
            Shape rect = targetSheet.Shapes.AddRectangle(2, 2, 100, 50, 0, 0);
            rect.Name = "DemoRectangle";

            // Verify that the shape belongs to the selected worksheet using Shape.Worksheet property
            Worksheet shapeSheet = rect.Worksheet;
            Console.WriteLine("Shape is on worksheet: " + shapeSheet.Name);

            // Save the workbook with the new shape (uses Workbook.Save(string) method)
            workbook.Save("output.xlsx");
        }
    }
}