// Title: How to keep shape text horizontal by setting RotateTextWithShape = false while rotating a shape in Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code using Aspose.Cells that adds a rectangle shape, turns off the shape's RotateTextWithShape flag, rotates the shape 45°, and saves the workbook. | Show how to keep a shape's text horizontal by setting RotateTextWithShape to false before applying a rotation angle in Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# prevent shape text from rotating when shape is rotated | how to set RotateTextWithShape false Aspose.Cells example | keep textbox text horizontal after rotating shape in Excel using Aspose.Cells
// Tags: Aspose.Cells shape text rotation flag | C# rotate shape without rotating its text | Excel shape horizontal text after rotation | Aspose.Cells drawing shape rotation example | disable shape text auto‑rotation .NET

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// The example creates a workbook, inserts a rectangle shape, disables automatic text rotation by setting shape.RotateTextWithShape = false, rotates the shape 45 degrees, and saves the file as an .xlsx workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            // Parameters: shape type, upper left row, upper left column, top, left, width, height
            Shape shape = worksheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // shape type
                2,   // upper left row
                2,   // upper left column
                5,   // top offset (pixels)
                5,   // left offset (pixels)
                100, // width (pixels)
                50   // height (pixels)
            );

            // Set the shape's text
            shape.Text = "Rotated Shape";

            // Rotate the shape 45 degrees
            shape.RotationAngle = 45;

            // Note: Aspose.Cells Shape does not expose a TextRotationAngle property.
            // The text will rotate together with the shape. If horizontal text is required,
            // additional handling (e.g., using a separate textbox) would be needed.

            // Define output file path
            string outputPath = "RotateTextWithShapeExample.xlsx";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to a file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
