using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class MoveWatermarkToBack
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Assume the watermark shape is on the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Iterate through all shapes to find the one named "Watermark"
        foreach (Shape shape in sheet.Shapes)
        {
            // The Shape class has a Name property that can be used to identify it
            if (shape.Name == "Watermark")
            {
                // Send the shape to the back layer (negative order)
                shape.ToFrontOrBack(-1);
                break; // Exit after moving the target shape
            }
        }

        // Save the modified workbook (replace with your desired output path)
        workbook.Save("output.xlsx");
    }
}