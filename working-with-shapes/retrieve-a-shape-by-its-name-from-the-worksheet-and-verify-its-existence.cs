using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class RetrieveShapeByName
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a rectangle shape and assign a custom name
        Shape shape = worksheet.Shapes.AddRectangle(1, 0, 0, 100, 100, 100);
        shape.Name = "MyRectangle";

        // Retrieve the shape by its name from the worksheet's shape collection
        Shape retrievedShape = worksheet.Shapes["MyRectangle"];

        // Verify whether the shape exists
        if (retrievedShape != null)
        {
            Console.WriteLine("Shape found: " + retrievedShape.Name);
        }
        else
        {
            Console.WriteLine("Shape not found.");
        }

        // Save the workbook (optional, demonstrates lifecycle usage)
        workbook.Save("RetrieveShapeByName.xlsx");
    }
}