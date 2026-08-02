// Title: Set a Uniform Adjustment Guide Value for All Shapes in Aspose.Cells (C#)
// Description: Creates a workbook, adds several auto‑shapes, then loops through every shape on the first worksheet. For each shape it accesses the ShapeGuideCollection via `shape.Geometry.ShapeAdjustValues`, assigns the same adjustment value (0.4) to all existing guides, and inserts a default guide named "adj" when a shape has none. The workbook is saved with the updated geometry.
// Keywords: Aspose.Cells shape adjustment | ShapeGuideCollection C# | batch update shape guides | uniform shape geometry Aspose | add default adjustment guide
// Common Searches: Aspose.Cells set same adjustment value for all shapes | loop through shapes and modify geometry guides C# | add missing adjustment guide to Aspose.Cells shape | how to standardize shape proportions in Aspose.Cells | C# example ShapeGuideCollection
// Developer Intent: Apply one adjustment value to every shape’s guides in a worksheet, adding a guide when none exist.
// Use Cases: Ensure all callout arrows have identical proportion across a financial report. | Give custom geometry shapes a consistent default guide before exporting to PDF. | Prepare a template workbook where shape sizes must stay uniform after bulk edits.
// AI Prompts: Write C# code that iterates over all shapes in an Aspose.Cells worksheet and sets each shape's adjustment guides to a specified uniform value, creating a default guide if the shape lacks any. | Explain the purpose of ShapeGuideCollection in Aspose.Cells and show how to read, modify, or add adjustment values for shapes inside a loop.

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;

// Creates a workbook, adds several auto‑shapes, then loops through every shape on the first worksheet. For each shape it accesses the ShapeGuideCollection via `shape.Geometry.ShapeAdjustValues`, assigns the same adjustment value (0.4) to all existing guides, and inserts a default guide named "adj" when a shape has none. The workbook is saved with the updated geometry.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add several shapes that support adjustment guides
        worksheet.Shapes.AddAutoShape(AutoShapeType.RightArrowCallout, 2, 0, 2, 0, 200, 150);
        worksheet.Shapes.AddAutoShape(AutoShapeType.Chevron, 5, 0, 5, 0, 150, 80);
        worksheet.Shapes.AddAutoShape(AutoShapeType.NotPrimitive, 8, 0, 8, 0, 120, 120); // custom geometry shape

        // Desired uniform adjustment value for all shapes
        double uniformValue = 0.4;

        // Iterate through each shape in the worksheet
        for (int i = 0; i < worksheet.Shapes.Count; i++)
        {
            Shape shape = worksheet.Shapes[i];

            // Get the collection of adjustment guides for the current shape
            ShapeGuideCollection guides = shape.Geometry.ShapeAdjustValues;

            if (guides.Count > 0)
            {
                // If the shape already has guides, set each one to the uniform value
                for (int g = 0; g < guides.Count; g++)
                {
                    guides[g].Value = uniformValue;
                }
            }
            else
            {
                // If no guides exist, add a default guide named "adj" with the uniform value
                guides.Add("adj", uniformValue);
            }
        }

        // Save the workbook with the modified shapes
        workbook.Save("UniformShapeAdjustments.xlsx");
    }
}
