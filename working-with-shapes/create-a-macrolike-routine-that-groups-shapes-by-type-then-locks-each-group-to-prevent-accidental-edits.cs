using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapeGroupingAndLocking
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add sample shapes of different types
            // Rectangles
            Shape rect1 = worksheet.Shapes.AddRectangle(2, 0, 2, 0, 60, 80);
            Shape rect2 = worksheet.Shapes.AddRectangle(5, 0, 5, 0, 60, 80);
            // Ovals
            Shape oval1 = worksheet.Shapes.AddOval(2, 0, 12, 0, 60, 80);
            Shape oval2 = worksheet.Shapes.AddOval(5, 0, 12, 0, 60, 80);
            // Text boxes
            Shape txt1 = worksheet.Shapes.AddTextBox(2, 0, 22, 0, 60, 80);
            Shape txt2 = worksheet.Shapes.AddTextBox(5, 0, 22, 0, 60, 80);

            // Group shapes by their AutoShapeType
            var groups = new Dictionary<AutoShapeType, List<Shape>>();

            foreach (Shape shape in worksheet.Shapes)
            {
                // Only consider AutoShape types (skip pictures, controls, etc.)
                if (shape.Type != null)
                {
                    AutoShapeType type = shape.Type;
                    if (!groups.ContainsKey(type))
                        groups[type] = new List<Shape>();
                    groups[type].Add(shape);
                }
            }

            // Iterate over each type group, create a group shape and lock it
            foreach (var kvp in groups)
            {
                List<Shape> shapeList = kvp.Value;
                // Need at least two shapes to form a meaningful group
                if (shapeList.Count < 2) continue;

                Shape[] shapesToGroup = shapeList.ToArray();

                // Group the shapes
                GroupShape groupShape = worksheet.Shapes.Group(shapesToGroup);

                // Lock the group to prevent accidental edits
                // Lock the group property itself
                groupShape.SetLockedProperty(ShapeLockType.Group, true);
                // Additionally, set the IsLocked flag (effective when worksheet is protected)
                groupShape.IsLocked = true;

                // Optional: give the group a descriptive name
                groupShape.Name = $"Group_{kvp.Key}";
            }

            // Save the workbook
            workbook.Save("GroupedAndLockedShapes.xlsx");
        }
    }
}