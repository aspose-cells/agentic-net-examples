// Title: Get absolute X/Y pixel coordinates of a named shape (Logo) in Aspose.Cells for .NET
// Description: Creates a workbook, adds a rectangle shape named "Logo" to the first worksheet, finds the shape by its Name, reads the shape's X and Y properties (pixel offsets from the worksheet's top‑left corner), writes the coordinates to the console for debugging, and saves the file.
// Keywords: Aspose.Cells | .NET | C# | shape coordinates | absolute position | X property | Y property | retrieve shape by name | worksheet shapes | debugging | pixel offsets | RetrieveShapeCoordinates | shape.X | shape.Y
// Common Searches: Aspose.Cells get shape X coordinate | How to read shape Y position in .NET | Retrieve absolute pixel location of a shape in Excel using Aspose.Cells | Log shape coordinates for debugging Aspose.Cells C# | Find shape named Logo and get its position
// Developer Intent: Read and output the absolute pixel X and Y positions of the shape named "Logo" in a worksheet.
// Use Cases: Diagnose layout issues by printing exact pixel locations of specific shapes. | Programmatically align or distribute shapes based on their absolute coordinates. | Include shape position data in logs or reports for audit trails.
// AI Prompts: Show C# code to obtain the X and Y pixel values of a shape called "Logo" with Aspose.Cells. | Give an example that iterates all worksheet shapes and logs each shape's absolute coordinates. | Explain how the X/Y properties of a shape map to the worksheet's pixel grid in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // Creates a workbook, adds a rectangle shape named "Logo" to the first worksheet, finds the shape by its Name, reads the shape's X and Y properties (pixel offsets from the worksheet's top‑left corner), writes the coordinates to the console for debugging, and saves the file.
    public class RetrieveShapeCoordinates
    {
        public static void Main()
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
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape and assign it the name "Logo"
            // Parameters: upperLeftRow, upperLeftColumn, upperLeftPixel,
            //             lowerRightRow, lowerRightColumn, lowerRightPixel
            Shape logoShape = worksheet.Shapes.AddRectangle(
                2,   // upperLeftRow (0‑based)
                2,   // upperLeftColumn (0‑based)
                0,   // upperLeftPixel
                5,   // lowerRightRow
                5,   // lowerRightColumn
                0);  // lowerRightPixel
            logoShape.Name = "Logo";

            // Retrieve the shape by its name (manual search)
            Shape retrievedShape = null;
            foreach (Shape shape in worksheet.Shapes)
            {
                if (shape.Name == "Logo")
                {
                    retrievedShape = shape;
                    break;
                }
            }

            if (retrievedShape != null)
            {
                // X and Y give the absolute offset from the worksheet's left/top borders (pixels)
                int absoluteX = retrievedShape.X;
                int absoluteY = retrievedShape.Y;

                // Log the coordinates for debugging
                Console.WriteLine($"Logo shape absolute X: {absoluteX} pixels");
                Console.WriteLine($"Logo shape absolute Y: {absoluteY} pixels");
            }
            else
            {
                Console.WriteLine("Shape named 'Logo' was not found.");
            }

            // Save the workbook (lifecycle: save)
            workbook.Save("RetrieveShapeCoordinates.xlsx");
        }
    }
}
