// Title: Adjust a Shape's Z‑Order by Adding 5 in Aspose.Cells for .NET (C#)
// Description: C# example that creates a workbook, adds overlapping rectangle shapes, reads the second shape's ZOrderPosition, adds five to it, assigns the new value, and saves the file to show the updated stacking order.
// Keywords: Aspose.Cells | C# shape ZOrderPosition | adjust shape Z-order | stacking order Excel | overlapping shapes Aspose | set ZOrderPosition | Excel shape layer | Aspose.Cells .NET example
// Common Searches: how to change shape Z-order Aspose.Cells | increase ZOrderPosition of a shape in C# | move shape forward in Excel using Aspose | set shape stacking order programmatically | Aspose.Cells shape layer example
// Developer Intent: Raise the Z‑order of a specific worksheet shape by five positions to change its visual layering.
// Use Cases: Bring annotation shapes to the front of overlapping charts in generated reports. | Place a watermark behind data cells after adding other graphics. | Define a custom layering sequence for dynamically added shapes in an Excel export.
// AI Prompts: Generate C# code with Aspose.Cells that increments a shape's ZOrderPosition by a given offset. | Show how to loop through all worksheet shapes and reorder them based on priority rules. | Provide a snippet that resets Z-order of shapes to a predefined sequence after creation.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// C# example that creates a workbook, adds overlapping rectangle shapes, reads the second shape's ZOrderPosition, adds five to it, assigns the new value, and saves the file to show the updated stacking order.
class AdjustShapeZOrder
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add two overlapping shapes to demonstrate Z-order changes
        Shape shape1 = sheet.Shapes.AddRectangle(5, 5, 100, 100, 0, 0);
        Shape shape2 = sheet.Shapes.AddRectangle(30, 30, 100, 100, 0, 0);

        // Display the initial Z-order position of the second shape
        Console.WriteLine("Initial ZOrderPosition of shape2: " + shape2.ZOrderPosition);

        // Increase the Z-order of shape2 by five positions
        int currentZ = shape2.ZOrderPosition;
        shape2.ZOrderPosition = currentZ + 5;

        // Display the new Z-order position to confirm the change
        Console.WriteLine("New ZOrderPosition of shape2: " + shape2.ZOrderPosition);

        // Save the workbook to observe the layer change in the resulting file
        workbook.Save("ShapeZOrderAdjusted.xlsx");
    }
}
