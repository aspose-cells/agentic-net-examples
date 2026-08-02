// Title: Group multiple shapes into a named GroupShape (IconSet) with Aspose.Cells for .NET (C#)
// Description: Shows how to add three rectangle shapes (Icon1‑Icon3), group them using ShapeCollection.Group, assign the group the name IconSet, and save the workbook as GroupedIcons.xlsx with Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | GroupShape | ShapeCollection.Group | group shapes | named shape group | Excel shape grouping | Aspose.Cells example | Rectangle shape | IconSet
// Common Searches: Aspose.Cells group shapes C# | How to create a GroupShape in Aspose.Cells | Name a GroupShape in Aspose.Cells | Combine multiple shapes into one in .NET Excel | ShapeCollection.Group method example
// Developer Intent: Combine several individual shapes into a single named group so they can be moved, resized, or formatted together.
// Use Cases: Move, resize, or rotate the entire IconSet with one operation. | Apply uniform formatting (fill, border, style) to all icons at once. | Toggle visibility or lock the whole group to simplify worksheet layout. | Copy or export the grouped icons as a single object. | Attach a hyperlink or macro to the group for interactive dashboards.
// AI Prompts: Generate C# code to ungroup the IconSet GroupShape and access each child shape. | Show how to add a hyperlink to a GroupShape named IconSet using Aspose.Cells. | Provide a loop that iterates through the shapes inside a GroupShape to change their fill color. | Explain how to copy a GroupShape to another worksheet in the same workbook. | Demonstrate setting the Z‑order of a GroupShape relative to other shapes.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapeGrouping
{
    // Shows how to add three rectangle shapes (Icon1‑Icon3), group them using ShapeCollection.Group, assign the group the name IconSet, and save the workbook as GroupedIcons.xlsx with Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Get the shape collection of the worksheet
            ShapeCollection shapes = sheet.Shapes;

            // Add three sample shapes and assign them the required names
            Shape icon1 = shapes.AddRectangle(2, 0, 2, 0, 40, 40);
            icon1.Name = "Icon1";

            Shape icon2 = shapes.AddRectangle(6, 0, 2, 0, 40, 40);
            icon2.Name = "Icon2";

            Shape icon3 = shapes.AddRectangle(10, 0, 2, 0, 40, 40);
            icon3.Name = "Icon3";

            // Prepare an array with the shapes to be grouped
            Shape[] iconsToGroup = new Shape[] { icon1, icon2, icon3 };

            // Group the shapes using ShapeCollection.Group (rule: ShapeCollection.Group)
            GroupShape iconSetGroup = shapes.Group(iconsToGroup);

            // Assign a name to the newly created group shape
            iconSetGroup.Name = "IconSet";

            // Save the workbook (lifecycle rule: save)
            workbook.Save("GroupedIcons.xlsx");
        }
    }
}
