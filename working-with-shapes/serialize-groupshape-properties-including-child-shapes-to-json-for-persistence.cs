// Title: How to serialize an Aspose.Cells GroupShape and its nested child shapes to JSON in C#
// AI Prompts: Generate C# code that walks through an Aspose.Cells GroupShape hierarchy, captures each shape's name, type, position, size, and rotation, builds a nested ShapeInfo object, and writes the result as indented JSON using System.Text.Json. | Create a recursive method in C# that extracts properties from a GroupShape and any inner GroupShapes in an Excel worksheet with Aspose.Cells, then saves the collected data to a JSON file.
// Common Searches: aspnet serialize group shape hierarchy to json using aspose.cells | c# extract child shapes from a group shape in an Excel file with Aspose.Cells | how to export Aspose.Cells group shape properties to a JSON file | recursive shape collection Aspose.Cells C# example | save Excel drawing objects as JSON with Aspose.Cells
// Tags: Aspose.Cells group shape JSON export | C# recursive group shape serialization Aspose.Cells | Excel drawing objects to JSON C# | System.Text.Json shape DTO Aspose.Cells | nested group shape extraction Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsGroupShapeSerialization
{
    // Simple DTO for a shape (including group shape)
    // The example loads an Excel workbook, locates the first GroupShape on the first worksheet, recursively extracts geometric and identification properties of the group and its child shapes into a ShapeInfo DTO hierarchy, serializes this hierarchy to a pretty‑printed JSON string with System.Text.Json, and writes the output to a file.
    public class ShapeInfo
    {
        public string Name { get; set; } = string.Empty;
        public string ShapeType { get; set; } = string.Empty;
        public double X { get; set; }          // Upper‑left X (points)
        public double Y { get; set; }          // Upper‑left Y (points)
        public double Width { get; set; }
        public double Height { get; set; }
        public double RotationAngle { get; set; }
        public bool IsGroup { get; set; }
        public List<ShapeInfo> ChildShapes { get; set; } = new List<ShapeInfo>();
    }

    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                const string inputPath = "input.xlsx";
                const string outputPath = "GroupShape.json";

                // Verify input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file '{inputPath}' not found.");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Assume the first worksheet contains the group shape
                Worksheet sheet = workbook.Worksheets[0];

                // Find the first GroupShape on the sheet
                GroupShape? groupShape = null;
                foreach (Shape shape in sheet.Shapes)
                {
                    if (shape is GroupShape gs)
                    {
                        groupShape = gs;
                        break;
                    }
                }

                if (groupShape == null)
                {
                    Console.WriteLine("No GroupShape found on the worksheet.");
                    return;
                }

                // Serialize the group shape and its children
                ShapeInfo groupInfo = SerializeGroupShape(groupShape);

                // Convert to JSON (pretty printed)
                var jsonOptions = new JsonSerializerOptions { WriteIndented = true };
                string json = JsonSerializer.Serialize(groupInfo, jsonOptions);

                // Persist JSON to a file
                File.WriteAllText(outputPath, json);

                Console.WriteLine($"GroupShape serialized to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Recursively extracts properties from a GroupShape and its child shapes
        private static ShapeInfo SerializeGroupShape(GroupShape group)
        {
            var info = new ShapeInfo
            {
                Name = group.Name ?? string.Empty,
                ShapeType = group.GetType().Name,
                X = group.X,
                Y = group.Y,
                Width = group.Width,
                Height = group.Height,
                RotationAngle = group.RotationAngle,
                IsGroup = true
            };

            // Use reflection to obtain the child shape collection (covers versions where the property may be missing)
            var childCollectionObj = group.GetType()
                                          .GetProperty("GroupShapeCollection")
                                          ?.GetValue(group) as ShapeCollection;

            if (childCollectionObj != null)
            {
                foreach (Shape child in childCollectionObj)
                {
                    var childInfo = new ShapeInfo
                    {
                        Name = child.Name ?? string.Empty,
                        ShapeType = child.GetType().Name,
                        X = child.X,
                        Y = child.Y,
                        Width = child.Width,
                        Height = child.Height,
                        RotationAngle = child.RotationAngle,
                        IsGroup = child is GroupShape
                    };

                    // If the child itself is a GroupShape, recurse
                    if (child is GroupShape nestedGroup)
                    {
                        childInfo.ChildShapes = SerializeGroupShape(nestedGroup).ChildShapes;
                    }

                    info.ChildShapes.Add(childInfo);
                }
            }

            return info;
        }
    }
}
