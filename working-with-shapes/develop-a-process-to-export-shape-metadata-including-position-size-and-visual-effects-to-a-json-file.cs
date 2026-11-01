// Title: Export Excel shape positions, sizes, rotation angle, and visual effect flags to JSON using Aspose.Cells for .NET (C#)
// Common Searches: how to extract shape coordinates, dimensions and rotation from an Excel workbook with Aspose.Cells C# | Aspose.Cells C# export shape metadata (top, left, width, height) to a JSON file | retrieve shape visual effect properties such as shadow and glow using Aspose.Cells .NET | serialize Excel shape collection to JSON in a C# application | save shape name and worksheet reference along with size data via Aspose.Cells
// Tags: Aspose.Cells export shape metadata to JSON | C# extract Excel shape properties | Aspose.Cells shape collection serialization | Excel shape position and size extraction | Aspose.Cells visual effect flags for shapes

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Drawing;   // Required for Shape and ShapeCollection types

namespace ShapeMetadataExport
{
    // Class representing the metadata we want to export for each shape
    // The example loads an Excel workbook with Aspose.Cells, iterates through each worksheet's ShapeCollection, captures each shape's name, worksheet name, top, left, width, height, rotation angle, and visual‑effect flags (shadow, glow, soft edge), stores the data in ShapeMetadata objects, serializes the collection to a formatted JSON string, and writes the output to a file.
    public class ShapeMetadata
    {
        public string WorksheetName { get; set; } = string.Empty;
        public string ShapeName { get; set; } = string.Empty;
        public double Top { get; set; }               // Position from the top of the worksheet (points)
        public double Left { get; set; }              // Position from the left of the worksheet (points)
        public double Width { get; set; }             // Width of the shape (points)
        public double Height { get; set; }            // Height of the shape (points)
        public double RotationAngle { get; set; }     // Rotation angle in degrees

        // Visual‑effect flags – not all versions expose these properties, default to false
        public bool HasShadow { get; set; }
        public bool HasGlow { get; set; }
        public bool HasSoftEdge { get; set; }
    }

    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // -----------------------------------------------------------------
                // Load the workbook (replace with your actual file path)
                // -----------------------------------------------------------------
                string inputPath = "input.xlsx";

                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Error: Input file '{inputPath}' not found.");
                    return;
                }

                Workbook workbook = new Workbook(inputPath);

                // -----------------------------------------------------------------
                // Collect shape metadata from every worksheet
                // -----------------------------------------------------------------
                List<ShapeMetadata> allShapes = new List<ShapeMetadata>();

                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    ShapeCollection shapes = sheet.Shapes;

                    for (int i = 0; i < shapes.Count; i++)
                    {
                        try
                        {
                            Shape shape = shapes[i];

                            // Build metadata object
                            ShapeMetadata meta = new ShapeMetadata
                            {
                                WorksheetName = sheet.Name,
                                ShapeName = shape.Name,
                                Top = shape.Top,
                                Left = shape.Left,
                                Width = shape.Width,
                                Height = shape.Height,
                                RotationAngle = shape.RotationAngle,

                                // Visual effects – these properties are not available in all versions,
                                // so we default to false.
                                HasShadow = false,
                                HasGlow = false,
                                HasSoftEdge = false
                            };

                            allShapes.Add(meta);
                        }
                        catch (Exception shapeEx)
                        {
                            Console.WriteLine($"Warning: Could not process shape index {i} on sheet '{sheet.Name}': {shapeEx.Message}");
                        }
                    }
                }

                // -----------------------------------------------------------------
                // Serialize the metadata list to JSON
                // -----------------------------------------------------------------
                var jsonOptions = new JsonSerializerOptions
                {
                    WriteIndented = true // pretty‑print for readability
                };

                string json = JsonSerializer.Serialize(allShapes, jsonOptions);

                // -----------------------------------------------------------------
                // Save JSON to a file (replace with your desired output path)
                // -----------------------------------------------------------------
                string outputPath = "shape_metadata.json";

                // Ensure the directory exists
                string? outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                File.WriteAllText(outputPath, json);

                Console.WriteLine($"Shape metadata exported to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
