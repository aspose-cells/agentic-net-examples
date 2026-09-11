// Title: Add a rectangle shape to cell B2 and rotate it 90° for diagonal orientation with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that inserts a rectangle shape at B2 in a new workbook and sets its RotationAngle to 90 degrees using Aspose.Cells. | Write a C# snippet to add a 150x100‑point rectangle shape to the first worksheet and rotate it diagonally with the Aspose.Cells drawing API. | Provide a C# example that modifies an existing Excel file to place a rectangle shape at row 2, column 2 and apply a 90‑degree rotation via Aspose.Cells.
// Common Searches: Aspose.Cells C# rotate shape 90 degrees in Excel | how to add and rotate a rectangle shape at B2 using Aspose.Cells .NET | C# Aspose.Cells set shape RotationAngle property example | draw diagonal rectangle in Excel with Aspose.Cells drawing API | programmatically rotate Excel shape with Aspose.Cells for .NET
// Tags: Aspose.Cells add rectangle shape C# | Aspose.Cells shape rotation angle | Aspose.Cells drawing API rotate shape | C# create rotated shape in Excel workbook | Excel shape diagonal orientation Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// // This program creates a new workbook, adds a rectangle shape at cell B2 (150 × 100 points), sets its RotationAngle to 90° for diagonal orientation, and saves the file as ShapeRotated.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle create)
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a rectangle shape at row 2, column 2 with width 150 and height 100 points
            // The AddShape method returns the created Shape object directly
            Shape shape = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // shape type
                1,   // upper‑left row index (zero‑based)
                1,   // upper‑left column index (zero‑based)
                0,   // upper‑left row offset in points
                0,   // upper‑left column offset in points
                150, // width in points
                100  // height in points
            );

            // Set rotation angle to 90 degrees for diagonal orientation
            shape.RotationAngle = 90;

            // Define output file name
            string outputPath = "ShapeRotated.xlsx";

            // Save the workbook (lifecycle save)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
