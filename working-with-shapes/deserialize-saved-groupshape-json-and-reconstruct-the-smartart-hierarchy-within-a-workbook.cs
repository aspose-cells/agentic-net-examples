// Title: Reconstruct SmartArt from GroupShape JSON with Aspose.Cells (C#)
// Description: This example shows how to read a JSON file that stores a GroupShape hierarchy, deserialize it into ShapeInfo objects, recursively create rectangle shapes on a worksheet, group child shapes, and finally assemble a root SmartArt group before saving the workbook.
// Keywords: Aspose.Cells | C# | GroupShape JSON | SmartArt reconstruction | deserialize shapes | recursive shape creation | Excel shape grouping | load shape hierarchy | JSON to Excel | Aspose.Cells example
// Common Searches: Aspose.Cells deserialize GroupShape JSON | How to rebuild SmartArt from JSON in C# | Create Excel shapes from JSON with Aspose.Cells | Recursive shape grouping Aspose.Cells example | Load saved SmartArt layout into workbook
// Developer Intent: Generate a SmartArt group in an Excel workbook by deserializing a saved GroupShape JSON file and rebuilding the shape hierarchy programmatically.
// Use Cases: Import a previously exported SmartArt design into a new workbook. | Synchronize complex diagram layouts across multiple Excel files using a JSON definition. | Build dynamic SmartArt diagrams from external data sources by first creating a JSON hierarchy.
// AI Prompts: Write a function that takes a GroupShape JSON string and returns a list of Aspose.Cells Shape objects preserving all group relationships. | Add comprehensive validation for row/column indices, dimensions, and null values in the recursive shape‑creation routine. | Generate sample JSON for a three‑level SmartArt diagram and show how the provided code reconstructs it in an Excel file.

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace SmartArtReconstruction
{
    // Classes that match the JSON structure of a saved GroupShape
    // This example shows how to read a JSON file that stores a GroupShape hierarchy, deserialize it into ShapeInfo objects, recursively create rectangle shapes on a worksheet, group child shapes, and finally assemble a root SmartArt group before saving the workbook.
    public class ShapeInfo
    {
        public string Name { get; set; }
        public string Text { get; set; }
        public int UpperLeftRow { get; set; }
        public int UpperLeftColumn { get; set; }
        public int Height { get; set; }          // in pixels
        public int Width { get; set; }           // in pixels
        public List<ShapeInfo> Children { get; set; } = new List<ShapeInfo>();
    }

    public class GroupShapeInfo
    {
        public List<ShapeInfo> Shapes { get; set; } = new List<ShapeInfo>();
    }

    public class Program
    {
        // Recursively creates shapes (or groups) based on the deserialized info
        private static List<Shape> CreateShapesRecursive(Worksheet ws, List<ShapeInfo> infos)
        {
            var createdShapes = new List<Shape>();

            foreach (var info in infos)
            {
                // Add a rectangle shape; offsets are set to 0
                Shape shape = ws.Shapes.AddRectangle(
                    info.UpperLeftRow,
                    info.UpperLeftColumn,
                    0,
                    0,
                    info.Width,
                    info.Height);

                shape.Name = info.Name;
                shape.Text = info.Text;

                // If the shape has child shapes, create them first and then group them
                if (info.Children != null && info.Children.Count > 0)
                {
                    // Create child shapes recursively
                    List<Shape> childShapes = CreateShapesRecursive(ws, info.Children);

                    // Group the child shapes into a new GroupShape
                    GroupShape childGroup = ws.Shapes.Group(childShapes.ToArray());
                    childGroup.Name = $"{info.Name}_Group";

                    // Add the group to the collection that will be returned for further grouping (if needed)
                    createdShapes.Add(childGroup);
                }
                else
                {
                    // No children – add the simple shape directly
                    createdShapes.Add(shape);
                }
            }

            return createdShapes;
        }

        public static void Main()
        {
            try
            {
                // 1. Create a new workbook (lifecycle rule: create)
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // 2. Load the JSON that represents a previously saved GroupShape
                string jsonPath = "groupShapeData.json";

                if (!File.Exists(jsonPath))
                {
                    Console.WriteLine($"JSON file not found: {jsonPath}");
                    return;
                }

                string jsonContent = File.ReadAllText(jsonPath);
                GroupShapeInfo groupInfo = JsonSerializer.Deserialize<GroupShapeInfo>(jsonContent);

                if (groupInfo?.Shapes == null || groupInfo.Shapes.Count == 0)
                {
                    Console.WriteLine("No shape information found in JSON.");
                    return;
                }

                // 3. Recreate the shapes hierarchy from the deserialized data
                List<Shape> topLevelShapes = CreateShapesRecursive(worksheet, groupInfo.Shapes);

                // 4. Group the top‑level shapes to form the root SmartArt group
                if (topLevelShapes.Count > 0)
                {
                    GroupShape rootGroup = worksheet.Shapes.Group(topLevelShapes.ToArray());
                    rootGroup.Name = "ReconstructedSmartArt";
                }

                // 5. Save the workbook (lifecycle rule: save)
                string outputPath = "ReconstructedSmartArt.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
