// Title: Check a shape's default RotationAngle and rotate it only when not already rotated with Aspose.Cells for .NET
// AI Prompts: Write C# code using Aspose.Cells that opens an existing Excel file or creates a new workbook, obtains the first shape on the first worksheet, reads its RotationAngle property, and sets the angle to a specified value only if the current RotationAngle equals the default (0). | Create a .NET example that adds a rectangle shape when the worksheet has no shapes, then conditionally applies a 45-degree rotation based on the shape's current RotationAngle using the Aspose.Cells API.
// Common Searches: asp.net aspose.cells check if shape rotation angle is default before setting | c# aspose.cells conditional shape rotation example | how to read shape RotationAngle property in Aspose.Cells | rotate Excel shape only when not already rotated using Aspose.Cells | default value of RotationAngle for shapes in Aspose.Cells .NET
// Tags: Aspose.Cells shape rotation angle check | C# conditional shape rotation Excel | Aspose.Cells add rectangle shape if missing | Aspose.Cells load or create workbook example | Aspose.Cells default shape rotation property

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an existing workbook or creates a new one, ensures a rectangle shape exists on the first worksheet, reads the shape's RotationAngle (default 0), and applies a 45-degree rotation only when the shape has not been rotated, then saves the workbook.
class RotateTextWithShapeExample
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            Workbook workbook;

            // Load existing workbook if it exists; otherwise create a new one
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                // Add a default worksheet
                workbook.Worksheets.Add("Sheet1");
            }

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure there is at least one shape; otherwise add a rectangle shape for demonstration
            Shape shape;
            if (worksheet.Shapes.Count > 0)
            {
                shape = worksheet.Shapes[0];
            }
            else
            {
                // Add a rectangle shape at position (row 2, column 2) with size 100x50 points
                shape = worksheet.Shapes.AddShape(MsoDrawingType.Rectangle, 2, 2, 0, 0, 100, 50);
            }

            // The RotationAngle property defines the shape's rotation in degrees.
            // Default value is 0 (no rotation).
            double currentAngle = shape.RotationAngle;

            Console.WriteLine("Current RotationAngle: " + currentAngle);

            // If the shape is not rotated, apply a rotation as an example.
            if (Math.Abs(currentAngle) < 0.001)
            {
                shape.RotationAngle = 45; // Rotate 45 degrees
                Console.WriteLine("RotationAngle set to 45 degrees.");
            }

            // Save the workbook to the desired output path
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
