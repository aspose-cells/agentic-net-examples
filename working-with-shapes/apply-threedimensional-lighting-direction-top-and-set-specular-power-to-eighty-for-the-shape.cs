using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            // Parameters: drawing type, upper left row, upper left column, upper left offset X, upper left offset Y, width, height
            Shape shape = worksheet.Shapes.AddShape(MsoDrawingType.Rectangle, 1, 1, 0, 0, 200, 100);

            // Access the shape's 3‑D formatting object
            ThreeDFormat threeDFormat = shape.ThreeDFormat;

            // Apply the required lighting direction
            threeDFormat.LightingDirection = LightRigDirectionType.Top;

            // Note: Aspose.Cells' ThreeDFormat does not expose a SpecularPower property.
            // If needed, other 3‑D properties such as Depth, RotationX/Y/Z can be set here.

            // Save the workbook with the applied 3‑D settings
            string outputPath = "Shape3DLightingSpecular.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}