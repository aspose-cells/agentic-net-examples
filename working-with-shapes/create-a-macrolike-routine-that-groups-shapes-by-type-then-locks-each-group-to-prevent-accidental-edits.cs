// Title: How to group shapes by drawing type and lock the groups in an Excel file with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# group shapes by MsoDrawingType and lock the group | protect grouped shapes in an Excel worksheet using Aspose.Cells .NET | C# code to create GroupShape objects and set IsLocked property with Aspose.Cells | lock shape groups after grouping shapes in Aspose.Cells workbook
// Tags: group shapes by MsoDrawingType Aspose.Cells | lock GroupShape objects Aspose.Cells C# | protect worksheet objects Aspose.Cells | create grouped shape collection Aspose.Cells | set IsLocked property for shapes Aspose.Cells

using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ShapeGroupingAndLockingDemo
{
    // The example creates a workbook, adds various shapes, groups shapes that share the same MsoDrawingType into GroupShape objects, marks each group as locked, protects the worksheet for objects, and saves the file as GroupedAndLockedShapes.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // -----------------------------------------------------------------
                // 1. Create a new workbook and get the first worksheet.
                // -----------------------------------------------------------------
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // -----------------------------------------------------------------
                // 2. Add a variety of shapes to the worksheet for demonstration.
                // -----------------------------------------------------------------
                // Rectangle
                Shape rect1 = sheet.Shapes.AddRectangle(2, 0, 2, 0, 80, 60);
                rect1.Name = "Rect1";

                Shape rect2 = sheet.Shapes.AddRectangle(6, 0, 2, 0, 80, 60);
                rect2.Name = "Rect2";

                // Oval
                Shape oval1 = sheet.Shapes.AddOval(10, 0, 2, 0, 70, 70);
                oval1.Name = "Oval1";

                Shape oval2 = sheet.Shapes.AddOval(14, 0, 2, 0, 70, 70);
                oval2.Name = "Oval2";

                // Line
                Shape line1 = sheet.Shapes.AddLine(18, 0, 2, 0, 100, 0);
                line1.Name = "Line1";

                Shape line2 = sheet.Shapes.AddLine(22, 0, 2, 0, 100, 0);
                line2.Name = "Line2";

                // TextBox
                Shape txt1 = sheet.Shapes.AddTextBox(26, 0, 2, 0, 120, 40);
                txt1.Name = "TextBox1";
                txt1.Text = "Sample TextBox";

                Shape txt2 = sheet.Shapes.AddTextBox(30, 0, 2, 0, 120, 40);
                txt2.Name = "TextBox2";
                txt2.Text = "Another TextBox";

                // -----------------------------------------------------------------
                // 3. Group shapes by their drawing type.
                // -----------------------------------------------------------------
                // Dictionary to collect shapes per type.
                Dictionary<MsoDrawingType, List<Shape>> shapesByType = new Dictionary<MsoDrawingType, List<Shape>>();

                // Iterate over all shapes in the worksheet.
                for (int i = 0; i < sheet.Shapes.Count; i++)
                {
                    Shape shp = sheet.Shapes[i];

                    // Use the MsoDrawingType property (not the AutoShapeType) for grouping.
                    MsoDrawingType type = shp.MsoDrawingType;

                    if (!shapesByType.ContainsKey(type))
                    {
                        shapesByType[type] = new List<Shape>();
                    }

                    shapesByType[type].Add(shp);
                }

                // List to keep references to created group shapes (so we can lock them).
                List<GroupShape> createdGroups = new List<GroupShape>();

                // Create a group for each type that has more than one shape.
                foreach (var kvp in shapesByType)
                {
                    List<Shape> list = kvp.Value;
                    if (list.Count < 2)
                        continue; // No need to group a single shape.

                    // Convert List<Shape> to Shape[] as required by Group method.
                    Shape[] shapeArray = list.ToArray();

                    // Group the shapes.
                    GroupShape group = sheet.Shapes.Group(shapeArray);
                    group.Name = $"Group_{kvp.Key}";
                    createdGroups.Add(group);
                }

                // -----------------------------------------------------------------
                // 4. Lock each created group to prevent accidental edits.
                // -----------------------------------------------------------------
                foreach (GroupShape group in createdGroups)
                {
                    // Setting IsLocked marks the shape as locked when the worksheet is protected.
                    group.IsLocked = true;
                }

                // -----------------------------------------------------------------
                // 5. Protect the worksheet for objects so that locked shapes cannot be edited.
                // -----------------------------------------------------------------
                sheet.Protect(ProtectionType.Objects);

                // -----------------------------------------------------------------
                // 6. Save the workbook.
                // -----------------------------------------------------------------
                string outFile = "GroupedAndLockedShapes.xlsx";
                workbook.Save(outFile);
                Console.WriteLine($"Workbook saved successfully to '{outFile}'.");
            }
            catch (Exception ex)
            {
                // Any unexpected error will be printed to the console.
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
