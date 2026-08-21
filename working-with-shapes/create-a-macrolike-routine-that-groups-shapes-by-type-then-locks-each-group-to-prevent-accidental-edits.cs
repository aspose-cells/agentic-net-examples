// Title: C# – Group Shapes by Type and Lock Them with Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, add rectangles, ovals and text boxes, collect shapes by their ShapeType, generate a GroupShape for each type that has multiple items, lock the group using ShapeLockType.Group, assign a meaningful name, and save the result as GroupedAndLockedShapes.xlsx.
// Keywords: Aspose.Cells group shapes C# | lock GroupShape Aspose.Cells | ShapeType grouping .NET | ShapeLockType.Group example | Aspose.Cells shape collection | C# workbook shape grouping | prevent shape editing Aspose
// Common Searches: how to group shapes by type in Aspose.Cells | lock a shape group in Aspose.Cells for .NET | Aspose.Cells create GroupShape from ShapeCollection | set ShapeLockType.Group property C# | Aspose.Cells shape grouping tutorial
// Developer Intent: Generate grouped Shape objects for each ShapeType and lock the groups to stop accidental modifications.
// Use Cases: Automatically bundle all rectangles on a worksheet into a locked group so users cannot move or resize them. | Combine every text box into a single locked group before sharing the workbook with clients. | Organize ovals into a protected group to preserve layout consistency during collaborative editing.
// AI Prompts: Write C# code using Aspose.Cells that iterates through a worksheet's ShapeCollection, groups shapes by Shape.Type, locks each GroupShape, and saves the workbook. | Show an example of assigning custom names to GroupShape objects created from shapes of the same type and then exporting the file.

using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapeGrouping
{
    // Demonstrates how to create a workbook, add rectangles, ovals and text boxes, collect shapes by their ShapeType, generate a GroupShape for each type that has multiple items, lock the group using ShapeLockType.Group, assign a meaningful name, and save the result as GroupedAndLockedShapes.xlsx.
    public class GroupAndLockShapesByType
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // ------------------------------------------------------------
                // Add sample shapes of different types for demonstration
                // ------------------------------------------------------------
                // Rectangles
                worksheet.Shapes.AddRectangle(2, 0, 2, 0, 60, 80);
                worksheet.Shapes.AddRectangle(5, 0, 2, 0, 60, 80);
                // Ovals
                worksheet.Shapes.AddOval(2, 0, 10, 0, 50, 50);
                worksheet.Shapes.AddOval(5, 0, 10, 0, 50, 50);
                // Text boxes
                worksheet.Shapes.AddTextBox(2, 0, 18, 0, 70, 30);
                worksheet.Shapes.AddTextBox(5, 0, 18, 0, 70, 30);
                // ------------------------------------------------------------

                // Get the shape collection of the worksheet
                ShapeCollection shapes = worksheet.Shapes;

                // Dictionary to collect shapes by their ShapeType (enum value)
                Dictionary<int, List<Shape>> shapesByType = new Dictionary<int, List<Shape>>();

                // Iterate through all shapes and group them by type
                for (int i = 0; i < shapes.Count; i++)
                {
                    Shape shape = shapes[i];
                    int typeKey = (int)shape.Type;

                    if (!shapesByType.ContainsKey(typeKey))
                    {
                        shapesByType[typeKey] = new List<Shape>();
                    }
                    shapesByType[typeKey].Add(shape);
                }

                // For each type that has more than one shape, create a group and lock it
                foreach (var kvp in shapesByType)
                {
                    List<Shape> shapeList = kvp.Value;
                    if (shapeList.Count < 2)
                        continue; // No need to group a single shape

                    // Convert the list to an array as required by the Group method
                    Shape[] shapesToGroup = shapeList.ToArray();

                    // Group the shapes
                    GroupShape groupShape = shapes.Group(shapesToGroup);

                    // Lock the group to prevent accidental edits
                    groupShape.SetLockedProperty(ShapeLockType.Group, true);

                    // Optional: give the group a meaningful name
                    groupShape.Name = $"Group_Type_{kvp.Key}";
                }

                // Save the workbook
                workbook.Save("GroupedAndLockedShapes.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            GroupAndLockShapesByType.Run();
        }
    }
}
