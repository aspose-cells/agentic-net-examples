// Title: Group and Lock Shapes by Type in Excel with Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, add rectangles, ovals and text boxes, automatically group shapes by their AutoShapeType, and apply lock settings (move, resize, selection) to each GroupShape before saving the file.
// Keywords: Aspose.Cells shape grouping | C# group shapes by type | lock Excel shapes Aspose.Cells | AutoShapeType grouping .NET | protect shape groups in Excel | Aspose.Cells GroupShape lock | Excel shape protection C#
// Common Searches: how to group shapes by type using Aspose.Cells | lock grouped shapes in Excel with Aspose.Cells .NET | Aspose.Cells example for protecting shape groups | C# code to prevent editing of Excel shapes | group and lock shapes in generated workbook
// Developer Intent: Create a reusable routine that collects shapes by their AutoShapeType, groups each set, and locks the groups to prevent accidental edits.
// Use Cases: Organize similar shapes (rectangles, ovals, text boxes) into separate locked groups in automated reports. | Distribute Excel templates where certain shape groups must remain static for end‑users. | Enforce shape protection in dashboards that contain interactive graphics.
// AI Prompts: Generate C# code using Aspose.Cells that groups all shapes by AutoShapeType and locks each group. | Show how to apply ShapeLockType flags (Move, Resize, Selection) to a GroupShape in Aspose.Cells. | Provide a complete example with error handling that saves a workbook after grouping and locking shapes.

using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapeGrouping
{
    // Entry point for the application
    // Demonstrates how to create a workbook, add rectangles, ovals and text boxes, automatically group shapes by their AutoShapeType, and apply lock settings (move, resize, selection) to each GroupShape before saving the file.
    public class Program
    {
        public static void Main(string[] args)
        {
            try
            {
                GroupAndLockShapesByType.Run();
                Console.WriteLine("Workbook 'GroupedAndLockedShapes.xlsx' created successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    public class GroupAndLockShapesByType
    {
        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add sample shapes of different types
            // Rectangle (AutoShapeType.Rectangle)
            Shape rect1 = worksheet.Shapes.AddRectangle(2, 0, 2, 0, 60, 80);
            Shape rect2 = worksheet.Shapes.AddRectangle(5, 0, 5, 0, 60, 80);
            // Oval (AutoShapeType.Ellipse)
            Shape oval1 = worksheet.Shapes.AddOval(2, 0, 12, 0, 60, 80);
            Shape oval2 = worksheet.Shapes.AddOval(5, 0, 12, 0, 60, 80);
            // TextBox (AutoShapeType.TextBox)
            Shape txt1 = worksheet.Shapes.AddTextBox(2, 0, 22, 0, 60, 80);
            Shape txt2 = worksheet.Shapes.AddTextBox(5, 0, 22, 0, 60, 80);

            // Collect shapes by their AutoShapeType
            Dictionary<AutoShapeType, List<Shape>> shapesByType = new Dictionary<AutoShapeType, List<Shape>>();
            foreach (Shape shape in worksheet.Shapes)
            {
                // Skip already grouped shapes (if any)
                if (shape.IsInGroup) continue;

                AutoShapeType type = shape.Type;
                if (!shapesByType.ContainsKey(type))
                {
                    shapesByType[type] = new List<Shape>();
                }
                shapesByType[type].Add(shape);
            }

            // For each type, group the shapes and lock the resulting group
            foreach (KeyValuePair<AutoShapeType, List<Shape>> entry in shapesByType)
            {
                List<Shape> shapeList = entry.Value;
                if (shapeList.Count == 0) continue;

                // Group the shapes of the same type
                Shape[] shapeArray = shapeList.ToArray();
                GroupShape group = worksheet.Shapes.Group(shapeArray);

                // Lock the group to prevent accidental edits
                group.SetLockedProperty(ShapeLockType.Group, true);
                group.SetLockedProperty(ShapeLockType.Move, true);
                group.SetLockedProperty(ShapeLockType.Resize, true);
                group.SetLockedProperty(ShapeLockType.Selection, true);
            }

            // Save the workbook
            workbook.Save("GroupedAndLockedShapes.xlsx");
        }
    }
}
