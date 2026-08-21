// Title: Serialize Aspose.Cells GroupShape and its child shapes to JSON (C#)
// Description: Creates a workbook, adds rectangle and oval shapes, groups them, maps the group and each child to DTO classes, serializes the model with System.Text.Json, writes the JSON file, and optionally saves the Excel file.
// Keywords: Aspose.Cells | GroupShape | C# | JSON serialization | shape hierarchy | export grouped shapes | persist Excel shapes | System.Text.Json | custom DTO | worksheet shapes
// Common Searches: Aspose.Cells serialize GroupShape to JSON | export grouped shapes as JSON C# | save Excel shape hierarchy JSON | convert GroupShape properties to JSON | Aspose.Cells group shape persistence
// Developer Intent: Capture all properties of a GroupShape and its child shapes and store them in a JSON file for later reuse.
// Use Cases: Recreate the exact layout of grouped shapes in another workbook by loading the JSON model. | Synchronize shape metadata across multiple Excel files through a shared JSON representation. | Produce an external report of shape types, dimensions, and positions for analytics or documentation.
// AI Prompts: Generate C# code that reads the saved groupShape.json, deserializes it into GroupShapeInfo, and rebuilds the GroupShape with its child shapes in a new worksheet. | Create a method to compare two GroupShapeInfo objects and output differences in their child ShapeInfo collections. | Write a utility that scans all worksheets for GroupShape objects, serializes each to a separate JSON file, and logs the generated file paths.

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;

// Create a new workbook and obtain the first worksheet
Workbook workbook = new Workbook();
Worksheet sheet = workbook.Worksheets[0];

// Add two sample shapes
Shape rect = sheet.Shapes.AddRectangle(0, 0, 0, 0, 100, 100);
rect.Name = "Rect1";
rect.AlternativeText = "First rectangle";

Shape oval = sheet.Shapes.AddOval(0, 0, 150, 0, 100, 100);
oval.Name = "Oval1";
oval.AlternativeText = "First oval";

// Group the shapes into a GroupShape
GroupShape group = sheet.Shapes.Group(new Shape[] { rect, oval });
group.Name = "MyGroup";
group.AlternativeText = "Group of shapes";

// Build a serializable model for the group shape
GroupShapeInfo groupInfo = new GroupShapeInfo
{
    Name = group.Name,
    AlternativeText = group.AlternativeText,
    Top = group.Top,
    Left = group.Left,
    Width = group.Width,
    Height = group.Height,
    ChildShapes = new List<ShapeInfo>()
};

// Populate child shape information
foreach (Shape child in group.GetGroupedShapes())
{
    groupInfo.ChildShapes.Add(new ShapeInfo
    {
        Name = child.Name,
        AlternativeText = child.AlternativeText,
        Type = child.Type.ToString(),
        Top = child.Top,
        Left = child.Left,
        Width = child.Width,
        Height = child.Height,
        IsGroup = child.IsGroup
    });
}

// Serialize the model to JSON
JsonSerializerOptions jsonOptions = new JsonSerializerOptions { WriteIndented = true };
string json = JsonSerializer.Serialize(groupInfo, jsonOptions);

// Persist JSON to a file
File.WriteAllText("groupShape.json", json);

// Save the workbook (optional persistence of the Excel file)
workbook.Save("GroupShapeDemo.xlsx");

// Serializable representation of an individual shape
// Creates a workbook, adds rectangle and oval shapes, groups them, maps the group and each child to DTO classes, serializes the model with System.Text.Json, writes the JSON file, and optionally saves the Excel file.
public class ShapeInfo
{
    public string Name { get; set; }
    public string AlternativeText { get; set; }
    public string Type { get; set; }
    public int Top { get; set; }
    public int Left { get; set; }
    public int Width { get; set; }
    public int Height { get; set; }
    public bool IsGroup { get; set; }
}

// Serializable representation of a group shape, including its child shapes
public class GroupShapeInfo
{
    public string Name { get; set; }
    public string AlternativeText { get; set; }
    public int Top { get; set; }
    public int Left { get; set; }
    public int Width { get; set; }
    public int Height { get; set; }
    public List<ShapeInfo> ChildShapes { get; set; }
}
