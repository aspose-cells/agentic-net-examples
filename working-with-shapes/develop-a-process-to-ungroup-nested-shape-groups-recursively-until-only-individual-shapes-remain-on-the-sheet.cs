using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapeUngroupDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (lifecycle rule: create)
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                ShapeCollection shapes = sheet.Shapes;

                // ---------- Create shapes ----------
                Shape rect1 = shapes.AddRectangle(2, 0, 2, 0, 80, 40);
                Shape rect2 = shapes.AddRectangle(6, 0, 2, 0, 80, 40);
                Shape oval1 = shapes.AddOval(10, 0, 2, 0, 80, 40);

                // Group all shapes together (nested groups are not supported directly)
                GroupShape group = shapes.Group(new Shape[] { rect1, rect2, oval1 });

                // Ungroup all groups recursively
                UngroupAllNestedGroups(sheet);

                // Save the workbook (lifecycle rule: save)
                workbook.Save("NestedShapesUngrouped.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        /// <summary>
        /// Recursively ungroups all group shapes on the specified worksheet
        /// until only individual (non‑group) shapes remain.
        /// </summary>
        /// <param name="worksheet">Worksheet containing the shapes.</param>
        static void UngroupAllNestedGroups(Worksheet worksheet)
        {
            ShapeCollection shapes = worksheet.Shapes;
            bool foundGroup;

            // Continue looping while at least one group shape is found in the collection
            do
            {
                foundGroup = false;

                // Iterate over the collection; note that the collection size may change after ungrouping
                for (int i = 0; i < shapes.Count; i++)
                {
                    Shape shape = shapes[i];

                    // Identify group shapes via the IsGroup property
                    if (shape.IsGroup)
                    {
                        // Cast to GroupShape and ungroup it (rule: GroupShape.Ungroup)
                        GroupShape group = (GroupShape)shape;
                        group.Ungroup();

                        // A group was ungrouped; restart the scan because the collection has changed
                        foundGroup = true;
                        break;
                    }
                }
            } while (foundGroup);
        }
    }
}