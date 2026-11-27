// Title: C# example: Deserialize a GroupShape JSON file and rebuild a SmartArt‑like hierarchy using rectangle shapes in Aspose.Cells
// Common Searches: how to read a GroupShape JSON file and create corresponding shapes with Aspose.Cells in C# | C# code to reconstruct SmartArt hierarchy from JSON using Aspose.Cells shapes | deserialize GroupShape JSON and add rectangle shapes programmatically in an Excel workbook | Aspose.Cells example for generating hierarchical shapes from JSON data | recursive shape creation from SmartArt node JSON in C# Aspose.Cells | save workbook after dynamically adding shapes based on JSON hierarchy Aspose.Cells
// Tags: Aspose.Cells JSON deserialization to shapes | C# recursive rectangle shape generation | SmartArt hierarchy reconstruction Aspose.Cells | GroupShape JSON mapping Aspose.Cells | dynamic shape positioning based on hierarchy level | saving workbook after shape insertion Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace SmartArtReconstruction
{
    // Classes that map to the JSON structure of a saved GroupShape
    // The sample reads a GroupShape.json file, deserializes it into C# objects, and recursively adds rectangle shapes to represent each SmartArt node. Shapes are positioned according to hierarchy depth using Aspose.Cells' Shapes.AddShape method, then the workbook is saved as an Excel file.
    public class GroupShapeJson
    {
        public string? Name { get; set; }
        public SmartArtJson? SmartArt { get; set; }
    }

    public class SmartArtJson
    {
        public string? Layout { get; set; }               // e.g., "BasicBlockList"
        public List<SmartArtNodeJson>? Nodes { get; set; }
    }

    public class SmartArtNodeJson
    {
        public string? Text { get; set; }
        public List<SmartArtNodeJson>? Children { get; set; }
    }

    class Program
    {
        static void Main()
        {
            try
            {
                // ---------- Load or create the workbook ----------
                Workbook workbook = new Workbook(); // creates a new workbook
                Worksheet sheet = workbook.Worksheets[0];

                // ---------- Read and deserialize the GroupShape JSON ----------
                string jsonPath = "GroupShape.json";

                if (!File.Exists(jsonPath))
                {
                    Console.WriteLine($"JSON file not found: {jsonPath}");
                    return;
                }

                string jsonContent = File.ReadAllText(jsonPath);
                GroupShapeJson? groupShapeData = JsonSerializer.Deserialize<GroupShapeJson>(jsonContent);

                if (groupShapeData?.SmartArt?.Nodes == null)
                {
                    Console.WriteLine("Invalid or empty SmartArt data in JSON.");
                    return;
                }

                // ---------- Reconstruct the SmartArt hierarchy ----------
                // NOTE: Aspose.Cells does not provide native SmartArt creation.
                // As a fallback, we create simple rectangle shapes for each node.
                // This demonstrates handling the JSON data without using unavailable APIs.

                int startX = 5;
                int startY = 5;
                int shapeWidth = 120;
                int shapeHeight = 30;
                int verticalSpacing = 10;

                // Recursive method to add shapes for each node
                void AddNodeShapes(List<SmartArtNodeJson> nodes, int level, ref int currentY)
                {
                    foreach (var node in nodes)
                    {
                        // Calculate position based on hierarchy level
                        int posX = startX + level * (shapeWidth + 20);
                        int posY = currentY;

                        // Add a rectangle shape representing the node
                        // Parameters: type, upperLeftRow, upperLeftColumn, rowOffset, columnOffset, height, width
                        Shape shape = sheet.Shapes.AddShape(
                            MsoDrawingType.Rectangle,
                            0,                     // upperLeftRow
                            0,                     // upperLeftColumn
                            posY,                  // rowOffset (pixels)
                            posX,                  // columnOffset (pixels)
                            shapeHeight,           // height (pixels)
                            shapeWidth);           // width (pixels)

                        shape.Text = node.Text ?? string.Empty;
                        shape.Name = $"Node_{Guid.NewGuid():N}";

                        // Update Y position for next shape
                        currentY += shapeHeight + verticalSpacing;

                        // Recurse into children, if any
                        if (node.Children != null && node.Children.Count > 0)
                        {
                            AddNodeShapes(node.Children, level + 1, ref currentY);
                        }
                    }
                }

                int currentYPos = startY;
                AddNodeShapes(groupShapeData.SmartArt.Nodes, 0, ref currentYPos);

                // ---------- Save the workbook ----------
                string outputPath = "ReconstructedSmartArt.xlsx";

                // Ensure the directory exists
                string? outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
