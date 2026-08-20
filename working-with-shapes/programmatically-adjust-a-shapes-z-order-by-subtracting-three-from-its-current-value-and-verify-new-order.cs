// Title: Decrease a Shape’s Z‑Order by 3 Using Aspose.Cells for .NET (C#)
// Description: Creates a workbook, adds a rectangle shape, reads its ZOrderPosition, subtracts three (clamped to zero), assigns the new value, prints both positions, and saves the file to confirm the change.
// Keywords: Aspose.Cells shape ZOrderPosition | C# adjust shape Z-order | lower shape layering Aspose.Cells | programmatic shape Z-order change | prevent negative Z-order Aspose.Cells
// Common Searches: how to lower shape Z-order by a specific amount in Aspose.Cells C# | set shape ZOrderPosition without negative value | Aspose.Cells move shape behind other objects | verify shape Z-order after modification .NET
// Developer Intent: Reduce a shape’s Z‑order by three positions and validate the updated order.
// Use Cases: Place a newly inserted graphic behind existing elements by decreasing its Z-order. | Maintain a back‑to‑front drawing order when adding shapes dynamically. | Reorder shapes based on custom priority while ensuring the Z-order never becomes negative.
// AI Prompts: Write C# code with Aspose.Cells that decreases a shape’s ZOrderPosition by a given offset and clamps the result at zero. | Show how to loop through all worksheet shapes and shift each Z-order down by three, preventing negative values. | Explain how to read, modify, and confirm a shape’s Z-order using Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a workbook, adds a rectangle shape, reads its ZOrderPosition, subtracts three (clamped to zero), assigns the new value, prints both positions, and saves the file to confirm the change.
class AdjustShapeZOrder
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a rectangle shape to the worksheet
        Shape shape = worksheet.Shapes.AddRectangle(10, 10, 100, 100, 0, 0);

        // Get the current Z-order position
        int currentZ = shape.ZOrderPosition;
        Console.WriteLine("Current ZOrderPosition: " + currentZ);

        // Subtract three from the current Z-order (ensure it does not become negative)
        int newZ = currentZ - 3;
        if (newZ < 0) newZ = 0;
        shape.ZOrderPosition = newZ;

        // Verify the new Z-order position
        Console.WriteLine("New ZOrderPosition after subtracting 3: " + shape.ZOrderPosition);

        // Save the workbook to verify changes
        workbook.Save("AdjustedZOrder.xlsx");
    }
}
