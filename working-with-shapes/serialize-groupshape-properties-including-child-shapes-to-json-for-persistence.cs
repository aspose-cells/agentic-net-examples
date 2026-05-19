using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace GroupShapeSerializationDemo
{
    // DTO for a generic shape (including group shapes)
    public class ShapeInfo
    {
        public string Name { get; set; }
        public string AlternativeText { get; set; }
        public int Top { get; set; }
        public int Left { get; set; }
        public int Height { get; set; }
        public int Width { get; set; }
        public bool IsGroup { get; set; }
        public List<ShapeInfo> Children { get; set; } = new List<ShapeInfo>();
    }

    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add two sample shapes
            Shape rect = sheet.Shapes.AddRectangle(0, 0, 0, 100, 100, 0);
            rect.Name = "MyRectangle";
            rect.AlternativeText = "Rect1";

            Shape oval = sheet.Shapes.AddOval(0, 100, 0, 100, 100, 0);
            oval.Name = "MyOval";
            oval.AlternativeText = "Oval1";

            // Group the shapes
            GroupShape group = sheet.Shapes.Group(new Shape[] { rect, oval });
            group.Name = "MyGroup";
            group.AlternativeText = "GroupOfTwo";

            // Build a serializable representation of the group shape
            ShapeInfo groupInfo = BuildShapeInfo(group);

            // Serialize to JSON
            string json = JsonSerializer.Serialize(groupInfo, new JsonSerializerOptions { WriteIndented = true });

            // Persist JSON to a file
            File.WriteAllText("GroupShape.json", json);

            // Optionally save the workbook to verify the shapes exist
            workbook.Save("GroupShapeDemo.xlsx");
        }

        // Recursively extracts shape properties; for group shapes it also processes child shapes
        static ShapeInfo BuildShapeInfo(Shape shape)
        {
            var info = new ShapeInfo
            {
                Name = shape.Name,
                AlternativeText = shape.AlternativeText,
                Top = shape.Top,
                Left = shape.Left,
                Height = shape.Height,
                Width = shape.Width,
                IsGroup = shape.IsGroup
            };

            // If the shape is a group, retrieve its child shapes
            if (shape.IsGroup)
            {
                GroupShape grp = (GroupShape)shape;
                foreach (Shape child in grp.GetGroupedShapes())
                {
                    info.Children.Add(BuildShapeInfo(child));
                }
            }

            return info;
        }
    }
}