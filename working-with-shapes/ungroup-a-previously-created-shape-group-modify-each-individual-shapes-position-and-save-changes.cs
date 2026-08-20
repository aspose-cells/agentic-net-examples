// Title: Ungroup Shapes, Adjust Their Positions, and Save the Workbook – Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook, add rectangle shapes, group them, ungroup the group, iterate through the individual shapes, modify each shape's Left and Top coordinates, and save the result as an XLSX file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# ungroup shape | move shape left top Aspose.Cells | modify shape position Excel .NET | group and ungroup shapes Aspose.Cells | Worksheet.Shapes iterate C# | Excel shape coordinates Aspose.Cells | UngroupAndModifyShapes example
// Common Searches: How to ungroup a shape group in Aspose.Cells C# | C# code to change Left and Top of Excel shapes after ungrouping | Aspose.Cells move individual shapes by index | Ungroup shapes and reposition them using Aspose.Cells for .NET | Adjust shape coordinates in an Excel workbook with Aspose.Cells
// Developer Intent: Separate a grouped shape into its original components, reposition each component programmatically, and persist the changes in a new Excel file.
// Use Cases: Create staggered or offset diagrams by ungrouping and shifting shapes after layout generation. | Programmatically correct overlapping shapes in dynamically generated reports. | Prepare workbooks for printing or PDF conversion by fine‑tuning shape positions post‑group manipulation.
// AI Prompts: Generate C# code that groups three shapes with Aspose.Cells, then ungroups them and sets each shape's Left and Top based on its order in the collection. | Show how to loop through Worksheet.Shapes in Aspose.Cells for .NET and move only non‑group shapes by a custom X/Y offset.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to create a workbook, add rectangle shapes, group them, ungroup the group, iterate through the individual shapes, modify each shape's Left and Top coordinates, and save the result as an XLSX file using Aspose.Cells for .NET.
class UngroupAndModifyShapes
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add two rectangle shapes to the worksheet
        Shape rect1 = worksheet.Shapes.AddRectangle(0, 0, 0, 0, 100, 50);
        Shape rect2 = worksheet.Shapes.AddRectangle(5, 0, 0, 0, 100, 50);

        // Group the two shapes
        GroupShape group = worksheet.Shapes.Group(new Shape[] { rect1, rect2 });

        // Ungroup the group shape
        group.Ungroup();

        // After ungrouping, the original shapes are back in the collection.
        // Modify each shape's position (Left and Top) individually.
        for (int i = 0; i < worksheet.Shapes.Count; i++)
        {
            Shape shape = worksheet.Shapes[i];
            // Skip any group shapes that might still exist
            if (shape.IsGroup) continue;

            // Example: shift each shape by 50 pixels to the right and 30 pixels down per index
            shape.Left += 50 * (i + 1);
            shape.Top  += 30 * (i + 1);
        }

        // Save the workbook with the modified shapes
        workbook.Save("UngroupedAndModifiedShapes.xlsx");
    }
}
