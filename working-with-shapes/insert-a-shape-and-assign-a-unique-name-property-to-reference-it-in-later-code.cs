using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapeDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            // Parameters: topRow, top (pixel offset), leftColumn, left (pixel offset), height, width
            Shape shape = worksheet.Shapes.AddRectangle(2, 0, 2, 0, 150, 200);

            // Assign a unique name to the shape for later reference
            shape.Name = "MyUniqueShape";

            // Example: retrieve the same shape later using its name
            Shape retrievedShape = worksheet.Shapes["MyUniqueShape"];
            if (retrievedShape != null)
            {
                // Modify a property to prove we have the correct shape
                retrievedShape.Text = "Hello from the named shape!";
            }

            // Save the workbook
            workbook.Save("ShapeWithUniqueName.xlsx");
        }
    }
}