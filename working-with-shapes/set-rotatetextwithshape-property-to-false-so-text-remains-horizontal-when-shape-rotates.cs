using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a text box shape to the worksheet
        Shape shape = worksheet.Shapes.AddTextBox(1, 0, 1, 0, 100, 200);
        shape.Text = "Sample Text";

        // Access the shape's text alignment settings
        ShapeTextAlignment textAlignment = shape.TextBody.TextAlignment;

        // Set RotateTextWithShape to false so the text stays horizontal when the shape rotates
        textAlignment.RotateTextWithShape = false;

        // Rotate the shape to demonstrate that the text remains horizontal
        shape.RotationAngle = 45;

        // Save the workbook
        workbook.Save("RotateTextWithoutShape.xlsx");
    }
}