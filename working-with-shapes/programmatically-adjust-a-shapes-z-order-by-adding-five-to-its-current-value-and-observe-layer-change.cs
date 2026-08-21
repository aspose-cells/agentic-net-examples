// Title: Adjust a Shape's Z‑Order by Adding 5 in Aspose.Cells for .NET
// Description: Creates a workbook, adds two overlapping rectangle shapes, sets their initial ZOrderPosition values, increments the second shape's ZOrderPosition by five, and saves the file so you can see the shape move forward in the stacking order.
// Keywords: Aspose.Cells .NET shape Z-order | C# ZOrderPosition | change shape stacking order Excel | move shape forward Aspose.Cells | adjust shape layer programmatically
// Common Searches: Aspose.Cells change shape Z-order C# | increase shape ZOrderPosition by 5 | bring a shape to front in Excel using Aspose | how to reorder overlapping shapes Aspose.Cells | set ZOrderPosition dynamically in .NET
// Developer Intent: Raise a specific shape's ZOrderPosition by five to place it ahead of other objects on the worksheet.
// Use Cases: Ensure a data label stays visible by moving it above chart bars. | Place a company logo in front of gridlines for clearer branding. | Push a watermark behind all content by decreasing its Z-order.
// AI Prompts: Generate C# code that adds a variable offset to a shape's ZOrderPosition while clamping the value within allowed limits. | Show how to move a shape backward by decreasing its ZOrderPosition and handle cases where the result would be negative. | Provide a method to list all shapes on a worksheet, retrieve their ZOrderPosition, and sort them from back to front.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a workbook, adds two overlapping rectangle shapes, sets their initial ZOrderPosition values, increments the second shape's ZOrderPosition by five, and saves the file so you can see the shape move forward in the stacking order.
class AdjustShapeZOrder
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add two overlapping rectangle shapes
        Shape shape1 = sheet.Shapes.AddRectangle(10, 10, 100, 100, 0, 0);
        Shape shape2 = sheet.Shapes.AddRectangle(30, 30, 100, 100, 0, 0);

        // Set initial Z-order positions (optional, for clarity)
        shape1.ZOrderPosition = 0;
        shape2.ZOrderPosition = 1;

        // Increase shape2's Z-order by 5 positions
        shape2.ZOrderPosition = shape2.ZOrderPosition + 5;

        // Save the workbook to observe the layer change
        workbook.Save("ShapeZOrderAdjusted.xlsx");
    }
}
