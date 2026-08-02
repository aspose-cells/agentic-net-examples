// Title: Aspose.Cells for .NET – Set Shape ZOrderPosition to Default Layer (0) and Verify
// Description: Creates a workbook, adds a rectangle shape, sets its ZOrderPosition to 0 (the base layer), confirms the value, and saves the file. Demonstrates how to control shape layering with Aspose.Cells in C#.
// Keywords: Aspose.Cells | C# | .NET | shape ZOrderPosition | default layer | Excel shape ordering | set shape Z-order | verify ZOrderPosition | worksheet shapes
// Common Searches: Aspose.Cells set shape ZOrderPosition to 0 | C# verify shape Z-order in Excel | how to move shape to bottom layer Aspose.Cells | default Z-order value for shapes in Aspose.Cells | check shape ZOrderPosition after setting
// Developer Intent: Programmatically assign a shape's ZOrderPosition to 0 to place it on the base layer and confirm the assignment using Aspose.Cells for .NET.
// Use Cases: Place a newly added chart behind all existing objects in a generated report. | Reset the layering of a shape after repositioning it to ensure it appears at the bottom. | Automated validation of shape order when building Excel workbooks dynamically.
// AI Prompts: Write C# code with Aspose.Cells that adds several shapes and sets one shape's ZOrderPosition to 0 so it appears at the bottom. | Explain how Aspose.Cells determines the default Z-order for worksheet shapes and how to retrieve the current ZOrderPosition. | Generate a C# unit test that asserts a shape's ZOrderPosition equals 0 after assignment.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsZOrderDemo
{
    // Creates a workbook, adds a rectangle shape, sets its ZOrderPosition to 0 (the base layer), confirms the value, and saves the file. Demonstrates how to control shape layering with Aspose.Cells in C#.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            // Parameters: upper left row, upper left column, top offset, left offset, height, width
            Shape shape = worksheet.Shapes.AddRectangle(5, 5, 0, 0, 100, 100);

            // Set the shape's Z-order position to zero (default layer position)
            shape.ZOrderPosition = 0;

            // Verify that the Z-order position is set to zero
            if (shape.ZOrderPosition == 0)
            {
                Console.WriteLine("Shape ZOrderPosition is correctly set to 0 (default layer).");
            }
            else
            {
                Console.WriteLine($"Unexpected ZOrderPosition: {shape.ZOrderPosition}");
            }

            // Save the workbook to a file
            workbook.Save("ShapeZOrderDemo.xlsx");
        }
    }
}
