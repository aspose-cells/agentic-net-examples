using System;
using System.Linq;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

class DisableRotateTextWithShape
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("Input.xlsx");

        // Get the first worksheet (adjust index as needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Select all shapes whose text rotation follows the shape
        var shapesWithRotation = worksheet.Shapes
            .Cast<Shape>()
            .Where(s => s.TextBody != null && s.TextBody.TextAlignment.RotateTextWithShape);

        // Disable rotation for each selected shape
        foreach (Shape shape in shapesWithRotation)
        {
            shape.TextBody.TextAlignment.RotateTextWithShape = false;
        }

        // Save the modified workbook
        workbook.Save("Output.xlsx");
    }
}