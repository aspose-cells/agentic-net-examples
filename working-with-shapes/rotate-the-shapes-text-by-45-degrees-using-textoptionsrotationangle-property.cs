// Title: Rotate Shape Text 45° with Aspose.Cells for .NET (C#)
// Description: C# sample that adds a rectangle shape to the first worksheet, assigns text, rotates the text 45 degrees using Shape.TextOptions.RotationAngle, and saves the workbook as an XLSX file with Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | rotate shape text | TextOptions.RotationAngle | shape text angle | Excel shape rotation | Aspose.Cells example | GitHub code sample | .NET workbook | 45 degree text rotation
// Common Searches: Aspose.Cells rotate shape text C# | Shape.TextOptions.RotationAngle example | How to rotate text inside a shape using Aspose.Cells | C# code to set text rotation angle in Excel shape | Aspose.Cells shape text 45 degrees
// Developer Intent: Apply a 45‑degree rotation to the text of a shape in an Excel workbook using Aspose.Cells for .NET.
// Use Cases: Create diagonal labels on forms or dashboards for better visual hierarchy. | Add angled watermarks or branding text to worksheets without affecting cell data. | Design custom callouts or annotations that need to align with chart trends.
// AI Prompts: Write C# code that rotates the text of all rectangle shapes in a workbook by a user‑specified angle using Aspose.Cells. | Explain the difference between Shape.RotationAngle and Shape.TextOptions.RotationAngle in Aspose.Cells. | Generate a step‑by‑step tutorial for dynamically rotating shape text based on runtime input in a .NET application.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// C# sample that adds a rectangle shape to the first worksheet, assigns text, rotates the text 45 degrees using Shape.TextOptions.RotationAngle, and saves the workbook as an XLSX file with Aspose.Cells for .NET.
class RotateShapeText
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
            Shape shape = worksheet.Shapes.AddRectangle(1, 0, 1, 0, 200, 100);
            shape.Text = "Rotated Text";

            // Rotate the shape (including its text) by 45 degrees
            shape.RotationAngle = 45;

            // Define output file path
            string outputPath = "ShapeTextRotated45.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
