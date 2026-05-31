using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class RotateShapeTextDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            // Parameters: upper left row, upper left column, top, left, width, height
            Shape shape = worksheet.Shapes.AddRectangle(1, 0, 1, 0, 150, 100);
            shape.Text = "Rotated Text";

            // Rotate the shape (including its text) by 45 degrees
            shape.RotationAngle = 45;

            // Save the workbook
            string outputPath = "ShapeTextRotated45.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}