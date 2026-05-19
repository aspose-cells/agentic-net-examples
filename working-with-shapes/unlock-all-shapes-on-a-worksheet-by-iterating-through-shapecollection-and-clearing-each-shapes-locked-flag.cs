using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class UnlockShapesDemo
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Get the collection of shapes on the current worksheet
            ShapeCollection shapes = sheet.Shapes;

            // Loop through each shape and unlock it
            for (int i = 0; i < shapes.Count; i++)
            {
                Shape shape = shapes[i];
                shape.IsLocked = false; // Unlock the shape
            }
        }

        // Save the workbook with the unlocked shapes (replace with desired output path)
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}