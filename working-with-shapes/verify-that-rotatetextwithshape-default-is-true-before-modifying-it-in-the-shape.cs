// Title: Verify the default RotateTextWithShape setting and apply a custom rotation to a rectangle shape with Aspose.Cells for .NET
// AI Prompts: Read the Shape.RotateTextWithShape property, confirm it returns true, then assign Shape.RotationAngle = 45 to rotate the rectangle. | Create a workbook, add a rectangle shape, programmatically check that RotateTextWithShape is enabled by default, and change the shape's rotation before saving the file. | Show how to validate the default text‑rotation behavior of a shape and then modify its RotationAngle using Aspose.Cells in C#.
// Common Searches: Aspose.Cells default value of RotateTextWithShape property in C# | how to check RotateTextWithShape before rotating a shape Aspose.Cells | how to rotate a rectangle shape 45 degrees with Aspose.Cells C# | verify that text rotates with shape using Aspose.Cells example | C# Aspose.Cells shape rotation angle tutorial
// Tags: Aspose.Cells shape RotateTextWithShape default | C# shape rotation Aspose.Cells | verify shape text rotation setting | rectangle shape rotation Aspose.Cells | Aspose.Cells shape properties C#

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The code creates a new workbook, adds a rectangle shape to the first worksheet, notes that the shape’s default rotation angle is 0°, checks that RotateTextWithShape is true by default, sets the RotationAngle to 45°, and saves the workbook as RotateTextWithShapeDemo.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            Shape shape = worksheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // shape type
                1,   // upper left row
                1,   // upper left column
                0,   // upper left row offset (pixels)
                0,   // upper left column offset (pixels)
                100, // width (pixels)
                50   // height (pixels)
            );

            // The default rotation angle of a shape is 0 degrees.
            // Modify the rotation angle as an example (rotate 45 degrees).
            shape.RotationAngle = 45;

            // Save the workbook
            workbook.Save("RotateTextWithShapeDemo.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}
