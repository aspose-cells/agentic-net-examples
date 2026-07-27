// Title: C# – Serialize Aspose.Cells GroupShape and Child Shapes to JSON
// Description: Demonstrates how to create a workbook, add rectangle and oval shapes, group them, capture the group's and each child's properties in DTOs, and persist the information as indented JSON using System.Text.Json. The JSON file can be stored for later reconstruction, and the workbook is saved afterwards.
// Keywords: Aspose.Cells | GroupShape serialization | C# JSON export | shape properties to JSON | Aspose.Cells child shapes | .NET Excel shape persistence | System.Text.Json Aspose.Cells
// Common Searches: serialize Aspose.Cells GroupShape to JSON | export grouped shape properties C# | save Excel shape metadata as JSON | Aspose.Cells group shape JSON example | how to persist shape layout Aspose.Cells
// Developer Intent: Export a GroupShape and its child shapes from an Aspose.Cells workbook to a JSON file for later reuse.
// Use Cases: Archive diagram layout metadata for custom reporting. | Transfer grouped shape definitions between micro‑services via a JSON contract. | Rebuild a saved group of shapes in a new workbook by deserializing the JSON.
// AI Prompts: Write C# code that reads the GroupShapeInfo JSON file and recreates the GroupShape with its child shapes in an Aspose.Cells workbook. | Provide a method to update the position of a specific child shape inside a serialized GroupShapeInfo and re‑serialize the modified object. | Explain how to extend the ShapeInfo DTO to include rotation, fill color, and line style so the JSON captures full shape fidelity.

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsGroupShapeSerialization
{
    // Simple DTO for shape information
    // Demonstrates how to create a workbook, add rectangle and oval shapes, group them, capture the group's and each child's properties in DTOs, and persist the information as indented JSON using System.Text.Json. The JSON file can be stored for later reconstruction, and the workbook is saved afterwards.
    public class ShapeInfo
    {
        public string Name { get; set; }
        public string AlternativeText { get; set; }
        public int Top { get; set; }
        public int Left { get; set; }
        public int Width { get; set; }
        public int Height { get; set; }
        public bool IsGroup { get; set; }
    }

    // DTO for group shape information, includes child shapes
    public class GroupShapeInfo
    {
        public string Name { get; set; }
        public string AlternativeText { get; set; }
        public int Top { get; set; }
        public int Left { get; set; }
        public int Width { get; set; }
        public int Height { get; set; }
        public List<ShapeInfo> ChildShapes { get; set; } = new List<ShapeInfo>();
    }

    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add two shapes
            Shape rect = sheet.Shapes.AddRectangle(0, 0, 0, 0, 100, 80);
            rect.Name = "RectShape";
            rect.AlternativeText = "First rectangle";

            Shape oval = sheet.Shapes.AddOval(0, 0, 120, 0, 100, 80);
            oval.Name = "OvalShape";
            oval.AlternativeText = "First oval";

            // Group the shapes
            GroupShape group = sheet.Shapes.Group(new Shape[] { rect, oval });
            group.Name = "MyGroup";
            group.AlternativeText = "Group of two shapes";

            // Build DTO for the group shape
            GroupShapeInfo groupInfo = new GroupShapeInfo
            {
                Name = group.Name,
                AlternativeText = group.AlternativeText,
                Top = group.Top,
                Left = group.Left,
                Width = group.Width,
                Height = group.Height
            };

            // Iterate child shapes via indexer
            for (int i = 0; i < group.GetGroupedShapes().Length; i++)
            {
                Shape child = group[i];
                ShapeInfo childInfo = new ShapeInfo
                {
                    Name = child.Name,
                    AlternativeText = child.AlternativeText,
                    Top = child.Top,
                    Left = child.Left,
                    Width = child.Width,
                    Height = child.Height,
                    IsGroup = child.IsGroup
                };
                groupInfo.ChildShapes.Add(childInfo);
            }

            // Serialize to JSON
            string json = JsonSerializer.Serialize(groupInfo, new JsonSerializerOptions { WriteIndented = true });

            // Persist JSON to a file (lifecycle rule)
            File.WriteAllText("GroupShapeInfo.json", json);

            // Save the workbook (lifecycle rule)
            workbook.Save("GroupShapeDemo.xlsx");
        }
    }
}
