using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class RetrieveShapesDemo
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook(); // lifecycle: create

        // Access the first worksheet (the selected worksheet)
        Worksheet worksheet = workbook.Worksheets[0];

        // Retrieve the Shapes collection from the worksheet
        ShapeCollection shapes = worksheet.Shapes; // property access

        // Example processing: list all shapes in the collection
        for (int i = 0; i < shapes.Count; i++)
        {
            Shape shape = shapes[i];
            Console.WriteLine($"Shape {i}: Type={shape.Type}, Name={shape.Name}");
        }

        // Save the workbook (optional, demonstrates lifecycle: save)
        workbook.Save("Output.xlsx");
    }
}