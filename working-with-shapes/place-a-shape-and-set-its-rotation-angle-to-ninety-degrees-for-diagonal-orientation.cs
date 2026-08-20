// Title: Rotate a Rectangle Shape 90° in Excel with Aspose.Cells for .NET
// Description: Shows how to create a workbook, insert a rectangle shape, set its RotationAngle to 90 degrees for a diagonal layout, optionally add text, and save the file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | rotate shape | RotationAngle | Excel rectangle | diagonal orientation | shape transformation | .NET | programmatic Excel | shape API
// Common Searches: Aspose.Cells rotate shape C# | set shape rotation angle Aspose.Cells | rotate rectangle in Excel using .NET | diagonal shape example Aspose.Cells | C# code rotate Excel shape 90 degrees
// Developer Intent: Programmatically rotate an Excel shape 90 degrees.
// Use Cases: Create a diagonal watermark for a report. | Design a tilted callout or arrow in a flowchart. | Add a rotated logo or badge to a spreadsheet. | Prepare a slanted background shape for printed forms.
// AI Prompts: Write C# code that adds multiple shapes to a worksheet and rotates each by a different angle using Aspose.Cells. | Show how to rotate a shape and align its text so the text follows the diagonal orientation. | Provide an example that rotates a shape and then adjusts its size to fit the rotated bounds. | Explain how to retrieve and modify the RotationAngle of an existing shape in a saved workbook.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // Shows how to create a workbook, insert a rectangle shape, set its RotationAngle to 90 degrees for a diagonal layout, optionally add text, and save the file using Aspose.Cells for .NET.
    public class ShapeDiagonalOrientationDemo
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            // Parameters: upper left row, upper left column, upper left row offset, upper left column offset, width, height
            Shape shape = worksheet.Shapes.AddRectangle(2, 2, 0, 0, 150, 100);

            // Set the rotation angle to 90 degrees for diagonal orientation
            shape.RotationAngle = 90;

            // Optionally set some text to visualize the rotation
            shape.Text = "Diagonal";

            // Define output file path
            string outputPath = "ShapeDiagonalOrientation.xlsx";

            // Ensure the directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook with the rotated shape
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
    }
}
