// Title: Ungroup Excel shapes with Aspose.Cells for .NET while preserving formatting
// Description: Loads a workbook, scans the worksheet's ShapeCollection in reverse, detects GroupShape objects, calls Ungroup() to separate them, and saves the file so each shape keeps its original size, position and style.
// Keywords: Aspose.Cells ungroup shapes | C# Excel shape grouping | GroupShape.Ungroup method | preserve shape formatting | reverse ShapeCollection iteration | Excel drawing objects .NET
// Common Searches: how to ungroup shapes in Excel using Aspose.Cells C# | preserve shape formatting when ungrouping Aspose.Cells | iterate ShapeCollection backwards Aspose.Cells | split grouped logo in Excel workbook .NET | remove shape groups programmatically Aspose
// Developer Intent: Separate every grouped shape in a workbook without altering the individual shapes' visual properties.
// Use Cases: Edit each component of a company logo that was saved as a single group. | Prepare a spreadsheet for a reporting tool that requires individual chart elements. | Clean up legacy workbooks by breaking down grouped drawings before applying per‑shape styling.
// AI Prompts: Generate C# code using Aspose.Cells to ungroup all group shapes in a worksheet and confirm that size, position, and style remain unchanged. | Explain why iterating a ShapeCollection in reverse is required when calling GroupShape.Ungroup() in Aspose.Cells. | Write a unit test that verifies shapes are ungrouped correctly and that their formatting persists after saving the workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Loads a workbook, scans the worksheet's ShapeCollection in reverse, detects GroupShape objects, calls Ungroup() to separate them, and saves the file so each shape keeps its original size, position and style.
class UngroupShapesDemo
{
    static void Main()
    {
        // Load a workbook that contains grouped shapes
        Workbook workbook = new Workbook("GroupedShapes.xlsx");
        Worksheet worksheet = workbook.Worksheets[0];
        ShapeCollection shapes = worksheet.Shapes;

        // Iterate in reverse because ungrouping modifies the collection
        for (int i = shapes.Count - 1; i >= 0; i--)
        {
            Shape shape = shapes[i];

            // Identify group shapes
            if (shape.IsGroup)
            {
                // Cast to GroupShape and ungroup; formatting of individual shapes is preserved
                GroupShape groupShape = (GroupShape)shape;
                groupShape.Ungroup();
            }
        }

        // Save the workbook with the shapes ungrouped
        workbook.Save("UngroupedShapes.xlsx");
    }
}
