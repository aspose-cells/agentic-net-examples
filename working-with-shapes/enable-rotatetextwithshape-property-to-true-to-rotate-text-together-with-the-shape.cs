using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

class RotateTextWithShapeDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a text box shape to the worksheet
        Shape shape = worksheet.Shapes.AddTextBox(1, 0, 1, 0, 150, 100);
        shape.Text = "Rotated with shape";

        // Enable rotating text together with the shape
        ShapeTextAlignment textAlignment = shape.TextBody.TextAlignment;
        textAlignment.RotateTextWithShape = true;

        // Rotate the shape to demonstrate the effect
        shape.RotationAngle = 45; // degrees

        // Save the workbook
        workbook.Save("RotateTextWithShapeDemo.xlsx");
    }
}