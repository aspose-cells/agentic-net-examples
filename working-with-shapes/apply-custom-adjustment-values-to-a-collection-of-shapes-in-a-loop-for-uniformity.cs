using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;

class ApplyUniformShapeAdjustments
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add shapes that support adjustment guides
        sheet.Shapes.AddAutoShape(AutoShapeType.RightArrowCallout, 2, 0, 2, 0, 200, 150);
        sheet.Shapes.AddAutoShape(AutoShapeType.Chevron, 5, 0, 5, 0, 150, 80);
        sheet.Shapes.AddAutoShape(AutoShapeType.NotPrimitive, 8, 0, 8, 0, 120, 120);

        // Desired uniform adjustment value for all guides
        double uniformValue = 0.3;

        // Loop through each shape in the worksheet
        for (int i = 0; i < sheet.Shapes.Count; i++)
        {
            Shape shape = sheet.Shapes[i];
            Geometry geometry = shape.Geometry;

            // Get the collection of adjustment guides for the shape
            ShapeGuideCollection guides = geometry.ShapeAdjustValues;

            // Apply the uniform value to every guide present
            for (int j = 0; j < guides.Count; j++)
            {
                guides[j].Value = uniformValue;
            }
        }

        // Save the modified workbook
        workbook.Save("UniformShapeAdjustments.xlsx");
    }
}