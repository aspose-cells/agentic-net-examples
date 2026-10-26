// Title: Create a rectangle shape and rotate it 90° for diagonal orientation in an Excel worksheet using Aspose.Cells for .NET (C#)
// Common Searches: how to rotate an Excel shape 90 degrees with Aspose.Cells in C# | Aspose.Cells add rectangle shape and set rotation angle programmatically | C# example for setting shape RotationAngle to achieve diagonal orientation in Excel | rotate worksheet shape to 90 degrees using Aspose.Cells .NET API | Aspose.Cells shape rotation sample code for rectangle
// Tags: Aspose.Cells Shapes.AddShape method | Shape.RotationAngle property usage | MsoDrawingType.Rectangle shape creation | Excel shape rotation with Aspose.Cells API | saving workbook after shape manipulation

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample creates a new workbook, adds a rectangle shape to the first worksheet, sets its RotationAngle to 90 degrees for a diagonal orientation, and saves the file as ShapeRotation.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            // Parameters: shape type, upper left row, upper left column, top offset, left offset, height, width
            Shape shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 2, 0, 2, 0, 100, 50);

            // Rotate the shape 90 degrees for diagonal orientation
            shape.RotationAngle = 90;

            // Define output file path
            string outputPath = "ShapeRotation.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
