using System;
using System.IO;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ShapePositioningDemo
{
    // Simple POCO to map configuration values
    public class ShapeConfig
    {
        public int X { get; set; }   // Horizontal offset in pixels
        public int Y { get; set; }   // Vertical offset in pixels
        public int Width { get; set; }   // Width in pixels
        public int Height { get; set; }  // Height in pixels
    }

    class Program
    {
        static void Main()
        {
            // Load configuration from a JSON file (e.g., shapeConfig.json)
            // Expected format: { "X": 150, "Y": 80, "Width": 200, "Height": 100 }
            string configPath = "shapeConfig.json";
            if (!File.Exists(configPath))
            {
                Console.WriteLine($"Configuration file '{configPath}' not found.");
                return;
            }

            string json = File.ReadAllText(configPath);
            ShapeConfig cfg = JsonSerializer.Deserialize<ShapeConfig>(json);

            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a rectangle shape; position will be adjusted later using X/Y
            // Parameters: upperLeftRow, upperLeftColumn, upperLeftRowOffset, upperLeftColumnOffset, width, height
            Shape shape = sheet.Shapes.AddRectangle(0, 0, 0, 0, cfg.Width, cfg.Height);

            // Set absolute position using pixel values from configuration
            shape.X = cfg.X;   // Horizontal offset from worksheet left border
            shape.Y = cfg.Y;   // Vertical offset from worksheet top border

            // Optionally set placement to MoveAndSize for predictable behavior
            shape.Placement = PlacementType.MoveAndSize;

            // Validation: output the set values and placement type
            Console.WriteLine($"Shape placed at X={shape.X} px, Y={shape.Y} px");
            Console.WriteLine($"Shape size: Width={shape.Width} px, Height={shape.Height} px");
            Console.WriteLine($"Placement type: {shape.Placement}");

            // Simple sanity check – ensure the shape's X/Y match the config
            bool isPositionValid = shape.X == cfg.X && shape.Y == cfg.Y;
            Console.WriteLine($"Position validation result: {(isPositionValid ? "PASS" : "FAIL")}");

            // Save the workbook to a file
            string outputPath = "ShapePositioned.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
    }
}