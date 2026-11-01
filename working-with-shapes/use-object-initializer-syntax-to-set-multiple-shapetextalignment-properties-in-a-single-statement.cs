// Title: Use a C# object initializer to configure Shape.TextHorizontalAlignment, Shape.TextVerticalAlignment, and RotationAngle in Aspose.Cells
// AI Prompts: Generate C# code that creates a rectangle shape in a workbook and assigns TextHorizontalAlignment, TextVerticalAlignment, and RotationAngle using a single object initializer expression with Aspose.Cells. | Refactor existing shape property assignments into an object initializer block for Aspose.Cells shape configuration in .NET.
// Common Searches: how to apply object initializer to set multiple shape properties in Aspose.Cells C# | Aspose.Cells set text horizontal and vertical alignment for shape in one line | C# object initializer for shape rotation angle Aspose.Cells example | initialize Aspose.Cells shape with TextHorizontalAlignment and RotationAngle using object initializer
// Tags: object initializer Aspose.Cells shape | shape text alignment Aspose.Cells | shape rotation angle Aspose.Cells | Aspose.Cells shape property initialization | C# shape configuration Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // Shows how to create a workbook, add a rectangle shape, and configure its TextHorizontalAlignment, TextVerticalAlignment, and RotationAngle in a single object initializer statement before saving the file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                var workbook = new Workbook();

                // Add a rectangle shape to the first worksheet
                Shape shape = workbook.Worksheets[0].Shapes.AddShape(
                    MsoDrawingType.Rectangle, // shape type
                    1,    // upper left row
                    1,    // upper left column
                    0,    // upper left row offset (in pixels)
                    0,    // upper left column offset (in pixels)
                    150,  // width (in pixels)
                    80);  // height (in pixels)

                // Set text alignment and rotation for the shape
                shape.TextHorizontalAlignment = TextAlignmentType.Center; // horizontal centering
                shape.TextVerticalAlignment = TextAlignmentType.Center;   // vertical centering
                shape.RotationAngle = 45;                                 // rotate text 45 degrees

                // Save the workbook
                workbook.Save("AlignedShape.xlsx");
                Console.WriteLine("Workbook saved successfully as AlignedShape.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
