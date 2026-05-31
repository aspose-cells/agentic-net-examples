using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class Program
{
    static void Main()
    {
        // Load an existing workbook
        Workbook workbook = new Workbook("input.xlsx");
        Worksheet sheet = workbook.Worksheets[0];

        // Locate the shape named "ChartOverlay"
        Shape overlayShape = null;
        foreach (Shape shape in sheet.Shapes)
        {
            if (shape.Name == "ChartOverlay")
            {
                overlayShape = shape;
                break;
            }
        }

        // If the shape is found, bring it to the front
        if (overlayShape != null)
        {
            // Positive value moves the shape forward in Z-order
            overlayShape.ToFrontOrBack(1);
        }

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}