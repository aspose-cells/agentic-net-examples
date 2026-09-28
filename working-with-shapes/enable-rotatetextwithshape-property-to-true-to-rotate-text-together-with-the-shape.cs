// Title: How to enable RotateTextWithShape and rotate a rectangle shape’s text in Aspose.Cells for .NET (C#)
// AI Prompts: Set shape.RotateTextWithShape = true, assign text, set shape.RotationAngle = 45, and save the workbook as an XLSX file using Aspose.Cells in C#. | Create a new Workbook, add a rectangle shape, enable text‑with‑shape rotation, rotate the shape, and export the result with Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# rotate shape text together with shape | Enable RotateTextWithShape property in Aspose.Cells example | Rotate rectangle shape and its text using Aspose.Cells for .NET | C# Aspose.Cells shape rotation angle with text rotation | Save rotated shape with text to XLSX using Aspose.Cells
// Tags: Aspose.Cells RotateTextWithShape API | C# shape rotation with text | Aspose.Cells rectangle shape example | export rotated shape to XLSX | text-with-shape rotation .NET

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a Workbook, adds a rectangle shape, sets its Text, enables the RotateTextWithShape property so the text rotates with the shape, applies a 45‑degree RotationAngle, and saves the file as RotateTextWithShape.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            // Parameters: shape type, upper left row, upper left column, top offset, left offset, height, width
            Shape shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 1, 1, 0, 0, 100, 50);

            // Set the text inside the shape
            shape.Text = "Rotated Text";

            // Rotate the shape (including its text) by 45 degrees
            shape.RotationAngle = 45;

            // Save the workbook to a file
            string outputPath = "RotateTextWithShape.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
