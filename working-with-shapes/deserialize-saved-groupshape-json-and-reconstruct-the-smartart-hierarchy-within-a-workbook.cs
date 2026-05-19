using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using System.Text.Json;

namespace AsposeCellsSmartArtReconstruction
{
    // JSON representation of a saved GroupShape
    public class ShapeInfo
    {
        public string? Type { get; set; }          // e.g., "Rectangle", "Oval"
        public double Left { get; set; }
        public double Top { get; set; }
        public double Width { get; set; }
        public double Height { get; set; }
        public string? Text { get; set; }
        public List<ShapeInfo> Children { get; set; } = new();
    }

    public class Program
    {
        // Recursively creates shapes from ShapeInfo and returns the created Shape objects
        private static List<Shape> CreateShapes(Worksheet worksheet, List<ShapeInfo>? infos)
        {
            var created = new List<Shape>();
            if (infos == null) return created;

            foreach (var info in infos)
            {
                try
                {
                    Shape shape = info.Type?.ToLower() switch
                    {
                        "oval" => worksheet.Shapes.AddOval(
                                    0, 0,
                                    (int)info.Top,
                                    (int)info.Left,
                                    (int)info.Height,
                                    (int)info.Width),
                        // default to rectangle
                        _ => worksheet.Shapes.AddRectangle(
                                    0, 0,
                                    (int)info.Top,
                                    (int)info.Left,
                                    (int)info.Height,
                                    (int)info.Width)
                    };

                    // Set common properties
                    shape.Left = (int)info.Left;
                    shape.Top = (int)info.Top;
                    shape.Width = (int)info.Width;
                    shape.Height = (int)info.Height;
                    if (!string.IsNullOrEmpty(info.Text))
                        shape.Text = info.Text;

                    created.Add(shape);

                    // Recursively create child shapes and group them
                    if (info.Children != null && info.Children.Count > 0)
                    {
                        var childShapes = CreateShapes(worksheet, info.Children);
                        if (childShapes.Count > 0)
                        {
                            GroupShape childGroup = worksheet.Shapes.Group(childShapes.ToArray());
                            childGroup.Name = $"{shape.Name}_Group";
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine($"Failed to create shape '{info.Type}': {ex.Message}");
                }
            }

            return created;
        }

        public static void Main()
        {
            try
            {
                // 1. Create a new workbook (or load an existing template if needed)
                Workbook workbook = new();
                Worksheet worksheet = workbook.Worksheets[0];

                // 2. Load the JSON that describes the saved GroupShape hierarchy
                string jsonPath = "groupshape.json";
                if (!File.Exists(jsonPath))
                    throw new FileNotFoundException($"JSON file not found: {jsonPath}");

                string jsonContent = File.ReadAllText(jsonPath);
                var rootShapes = JsonSerializer.Deserialize<List<ShapeInfo>>(jsonContent);
                if (rootShapes == null)
                    throw new InvalidOperationException("Failed to deserialize shape information.");

                // 3. Recreate the shapes in the worksheet
                List<Shape> topLevelShapes = CreateShapes(worksheet, rootShapes);

                // 4. Group top‑level shapes to mimic the original hierarchy
                GroupShape? smartArtGroup = null;
                if (topLevelShapes.Count > 1)
                {
                    smartArtGroup = worksheet.Shapes.Group(topLevelShapes.ToArray());
                    smartArtGroup.Name = "ReconstructedSmartArtGroup";
                }
                else if (topLevelShapes.Count == 1 && topLevelShapes[0] is GroupShape gs)
                {
                    smartArtGroup = gs;
                }

                // 5. Optional: attempt to retrieve SmartArt result (will be null if not SmartArt)
                if (smartArtGroup != null)
                {
                    GroupShape? result = smartArtGroup.GetResultOfSmartArt();
                    if (result != null)
                    {
                        result.Left += 50;
                        result.Top += 20;
                    }
                }

                // 6. Save the workbook with the reconstructed SmartArt hierarchy
                string outputPath = "ReconstructedSmartArt.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}