// Title: How to deserialize a GroupShape JSON file and rebuild a SmartArt hierarchy in an Aspose.Cells workbook using C#
// AI Prompts: Write C# code that reads a GroupShape JSON file, parses it into objects, and adds a SmartArt shape with the specified layout, position, and size to an Aspose.Cells worksheet. | Create a C# method that converts a layout name string from the JSON into the Aspose.Cells.Drawing.SmartArtLayout enum using reflection. | Implement a recursive C# function that takes a list of NodeInfo objects and populates the corresponding SmartArt nodes and their child nodes in an Aspose.Cells workbook.
// Common Searches: c# Aspose.Cells deserialize GroupShape JSON to SmartArt | how to add SmartArt from JSON definition in Aspose.Cells | using reflection to get SmartArtLayout enum in Aspose.Cells C# | recursive population of SmartArt nodes from custom objects Aspose.Cells
// Tags: deserialize GroupShape JSON Aspose.Cells | rebuild SmartArt hierarchy C# | dynamic SmartArt layout resolution reflection | recursive SmartArt node population Aspose.Cells | add SmartArt shape from custom JSON Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace SmartArtReconstruction
{
    // Classes that map to the JSON structure of the saved GroupShape
    // The program loads a GroupShape definition from a JSON file, deserializes it into C# objects, dynamically resolves the SmartArt layout, and recursively creates SmartArt nodes to reconstruct the hierarchy in a new Aspose.Cells workbook, which is then saved.
    public class GroupShape
    {
        public List<ShapeInfo> Shapes { get; set; }
    }

    public class ShapeInfo
    {
        public string Type { get; set; }               // e.g., "SmartArt"
        public SmartArtInfo SmartArt { get; set; }     // present when Type == "SmartArt"
        public PositionInfo Position { get; set; }     // cell location and offsets
        public SizeInfo Size { get; set; }             // width & height in points
    }

    public class SmartArtInfo
    {
        public string Layout { get; set; }                     // e.g., "BasicCycle"
        public List<NodeInfo> Nodes { get; set; }              // hierarchical node data
    }

    public class NodeInfo
    {
        public string Text { get; set; }
        public List<NodeInfo> Children { get; set; }
    }

    public class PositionInfo
    {
        public int Row { get; set; }
        public int Column { get; set; }
        public int OffsetX { get; set; }   // in points
        public int OffsetY { get; set; }   // in points
    }

    public class SizeInfo
    {
        public double Width { get; set; }   // in points
        public double Height { get; set; }  // in points
    }

    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Load JSON definition of the GroupShape
                string jsonPath = "groupshape.json";
                GroupShape groupShape = LoadGroupShape(jsonPath);

                // Reconstruct each shape in the workbook
                foreach (var shapeInfo in groupShape.Shapes ?? new List<ShapeInfo>())
                {
                    if (shapeInfo.Type != null &&
                        shapeInfo.Type.Equals("SmartArt", StringComparison.OrdinalIgnoreCase) &&
                        shapeInfo.SmartArt != null)
                    {
                        // Resolve layout enum dynamically
                        dynamic layout = GetSmartArtLayout(shapeInfo.SmartArt.Layout);

                        // Use dynamic invocation to avoid compile‑time dependency on SmartArt types
                        dynamic shapes = sheet.Shapes;
                        dynamic smartArt = shapes.AddSmartArt(
                            shapeInfo.Position.Row,
                            shapeInfo.Position.Column,
                            shapeInfo.Position.OffsetX,
                            shapeInfo.Position.OffsetY,
                            shapeInfo.Size.Width,
                            shapeInfo.Size.Height,
                            layout);

                        // Populate the SmartArt nodes recursively
                        PopulateSmartArtNodes(smartArt, shapeInfo.SmartArt.Nodes);
                    }
                    // Additional shape types (e.g., pictures, charts) could be handled here
                }

                // Save the workbook with the reconstructed SmartArt hierarchy
                string outputPath = "output.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Safely loads the JSON file and deserializes it
        private static GroupShape LoadGroupShape(string path)
        {
            try
            {
                if (File.Exists(path))
                {
                    string jsonContent = File.ReadAllText(path);
                    return JsonSerializer.Deserialize<GroupShape>(jsonContent) ?? new GroupShape { Shapes = new List<ShapeInfo>() };
                }
                Console.WriteLine($"JSON file not found: {path}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load JSON: {ex.Message}");
            }
            return new GroupShape { Shapes = new List<ShapeInfo>() };
        }

        // Retrieves the SmartArtLayout enum value via reflection; defaults to BasicCycle
        private static dynamic GetSmartArtLayout(string layoutName)
        {
            const string enumFullName = "Aspose.Cells.Drawing.SmartArtLayout, Aspose.Cells";
            Type enumType = Type.GetType(enumFullName);
            if (enumType != null)
            {
                if (!string.IsNullOrWhiteSpace(layoutName) &&
                    Enum.TryParse(enumType, layoutName, true, out var parsed))
                {
                    return parsed;
                }
                // Fallback to BasicCycle if available
                if (Enum.IsDefined(enumType, "BasicCycle"))
                {
                    return Enum.Parse(enumType, "BasicCycle");
                }
            }
            // If the enum type cannot be resolved, return null (AddSmartArt will throw at runtime)
            return null;
        }

        // Recursively adds nodes and their children to a SmartArt object using dynamic typing
        private static void PopulateSmartArtNodes(dynamic smartArt, List<NodeInfo> nodes)
        {
            if (nodes == null || nodes.Count == 0)
                return;

            foreach (var nodeInfo in nodes)
            {
                dynamic node = smartArt.Nodes.AddNode();   // add a new root node
                node.Text = nodeInfo.Text ?? string.Empty;
                AddChildNodes(node, nodeInfo.Children);
            }
        }

        // Helper method to add child nodes to a given parent node using dynamic typing
        private static void AddChildNodes(dynamic parentNode, List<NodeInfo> children)
        {
            if (children == null || children.Count == 0)
                return;

            foreach (var childInfo in children)
            {
                dynamic childNode = parentNode.Children.AddNode();
                childNode.Text = childInfo.Text ?? string.Empty;
                AddChildNodes(childNode, childInfo.Children);
            }
        }
    }
}
