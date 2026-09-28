// Title: Rotate rectangle shape text 45 degrees with Aspose.Cells for .NET (C#)
// AI Prompts: Create an Excel workbook in C# and add a rectangle shape whose text is rotated 45 degrees using the Shape.RotationAngle property. | Write C# code that inserts a rectangle shape, sets its Text property, and applies a 45‑degree rotation to the shape’s text with Aspose.Cells. | Generate a .xlsx file containing a rotated‑text rectangle shape by setting Shape.RotationAngle to 45 in Aspose.Cells for .NET.
// Common Searches: C# Aspose.Cells rotate shape text 45 degrees | How to set RotationAngle for a shape in Aspose.Cells .NET | Aspose.Cells example rotating rectangle shape text | Set shape text rotation angle in Excel using Aspose.Cells | Rotate text inside a shape with Aspose.Cells for .NET
// Tags: Aspose.Cells shape RotationAngle property | C# add rectangle shape to worksheet | Excel shape text rotation example | Aspose.Cells set shape text | rotate shape text Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace RotatedShapeExample
{
    // Demonstrates how to create a workbook, insert a rectangle shape, assign text, rotate the shape (and its text) 45 degrees via the Shape.RotationAngle property, and save the file as RotatedShape.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a rectangle shape to the worksheet
                // The AddShape method returns the created Shape object
                Shape shape = worksheet.Shapes.AddShape(
                    MsoDrawingType.Rectangle, // shape type
                    1,    // upper‑left row
                    0,    // upper‑left column
                    0,    // upper‑left row offset (in pixels)
                    100,  // upper‑left column offset (in pixels)
                    100,  // height (in pixels)
                    100   // width (in pixels)
                );

                // Set the shape's text
                shape.Text = "Rotated Text";

                // Rotate the shape (including its text) by 45 degrees
                shape.RotationAngle = 45;

                // Define output file path
                string outputPath = "RotatedShape.xlsx";

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
}
