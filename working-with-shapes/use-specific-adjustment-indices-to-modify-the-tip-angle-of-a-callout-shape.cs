// Title: C# – Adjust Callout Shape Tip Angle via ShapeGuide Indices in Aspose.Cells
// Description: Creates a workbook, inserts a Right Arrow Callout auto‑shape, accesses its ShapeGuideCollection, sets the first guide (index 0) to change the tip angle, optionally modifies a second guide, and saves the file.
// Keywords: Aspose.Cells callout tip angle | ShapeGuideCollection C# | adjust shape guide index Aspose.Cells | modify callout geometry .NET | auto shape adjustment Aspose
// Common Searches: how to change callout tip angle Aspose.Cells | shape guide index for callout tip Aspose | set callout arrow angle C# Aspose.Cells | adjust multiple shape guides Aspose.Cells example
// Developer Intent: Set specific ShapeGuide values to control the tip angle (and other parameters) of a callout auto‑shape.
// Use Cases: Create Excel diagrams with callouts that have a custom arrow tip angle. | Programmatically fine‑tune several adjustment guides for precise visual styling. | Generate reports where callout arrows need non‑default angles for clearer data annotation.
// AI Prompts: Show C# code that enumerates all ShapeGuide indices and their current values for a callout shape using Aspose.Cells. | Write code to set the tip angle of a callout to 30% of its maximum and the second guide to 10% in Aspose.Cells for .NET. | Explain how the valid value range for ShapeGuide adjustments varies across different callout types in Aspose.Cells.

using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a workbook, inserts a Right Arrow Callout auto‑shape, accesses its ShapeGuideCollection, sets the first guide (index 0) to change the tip angle, optionally modifies a second guide, and saves the file.
class ModifyCalloutTipAngle
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a callout shape (Right Arrow Callout) to the worksheet
        // Parameters: shape type, upper left row, upper left column, top offset, left offset, height, width
        Shape callout = worksheet.Shapes.AddAutoShape(
            AutoShapeType.RightArrowCallout, 2, 0, 2, 0, 150, 200);

        // Access the collection of shape adjustment guides
        ShapeGuideCollection guides = callout.Geometry.ShapeAdjustValues;

        // Modify the tip angle using specific adjustment indices
        // For many callout shapes, the first guide (index 0) controls the tip angle
        if (guides.Count > 0)
        {
            // Set the tip angle adjust value (range depends on shape, typically 0.0 – 1.0)
            guides[0].Value = 0.45; // Example: 45% of the maximum tip angle
        }

        // Optionally modify additional adjustment guides by index if needed
        // Example: adjust the second guide (index 1) if it exists
        if (guides.Count > 1)
        {
            guides[1].Value = 0.25;
        }

        // Save the workbook with the modified shape
        workbook.Save("CalloutTipAngleModified.xlsx");
    }
}
