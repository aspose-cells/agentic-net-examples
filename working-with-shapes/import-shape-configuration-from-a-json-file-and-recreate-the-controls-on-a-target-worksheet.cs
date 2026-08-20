// Title: C# – Import Shape Settings from JSON and Recreate Shapes in an Aspose.Cells Worksheet
// Description: Loads a JSON file that defines shape type, position, size, name, linked cell and input range, deserializes it into ShapeConfig objects, and uses Aspose.Cells to add rectangles, lines, ovals or generic shapes to the first worksheet of a new workbook. Missing dimensions are estimated from cell spans, optional properties are applied, and the workbook is saved as an XLSX file.
// Keywords: Aspose.Cells JSON shape import | C# add shapes from configuration | Aspose.Cells create rectangle line oval | linked cell shape Aspose.Cells | shape input range Aspose.Cells | deserialize shape settings C# | dynamic shape generation Aspose.Cells
// Common Searches: How to import shape definitions from JSON into Aspose.Cells | C# add rectangle, line, oval to worksheet using Aspose.Cells | Set linked cell for a shape with Aspose.Cells API | Calculate shape size from cell range in Aspose.Cells | Fallback shape type handling Aspose.Cells
// Developer Intent: Read a JSON file containing shape parameters and programmatically add matching shapes to an Aspose.Cells worksheet.
// Use Cases: Load shape definitions from a JSON file and place rectangles, lines, or ovals on the first worksheet with correct row/column anchors. | Provide explicit Width and Height in the JSON to override automatic size estimation. | Gracefully handle unknown shape types by defaulting to a generic rectangle shape.
// AI Prompts: Generate C# code that reads a JSON array of shape configurations and uses Aspose.Cells to add each shape to a worksheet, estimating missing dimensions from cell spans. | Create a JSON schema for the ShapeConfig class compatible with the provided Aspose.Cells shape‑creation logic. | Show how to extend the switch expression to support additional shape types such as "Triangle" using Aspose.Cells MsoDrawingType.

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ShapeImportExample
{
    // Represents the configuration for a shape read from JSON
    // Loads a JSON file that defines shape type, position, size, name, linked cell and input range, deserializes it into ShapeConfig objects, and uses Aspose.Cells to add rectangles, lines, ovals or generic shapes to the first worksheet of a new workbook. Missing dimensions are estimated from cell spans, optional properties are applied, and the workbook is saved as an XLSX file.
    public class ShapeConfig
    {
        public string? Type { get; set; }               // e.g., "Rectangle", "Line", "Oval"
        public string? Name { get; set; }               // Shape name
        public int UpperLeftRow { get; set; }           // Starting row (0‑based)
        public int UpperLeftColumn { get; set; }        // Starting column (0‑based)
        public int LowerRightRow { get; set; }          // Ending row (0‑based)
        public int LowerRightColumn { get; set; }       // Ending column (0‑based)
        public int Width { get; set; }                  // Width in pixels (optional)
        public int Height { get; set; }                 // Height in pixels (optional)
        public string? LinkedCell { get; set; }         // e.g., "A1"
        public string? InputRange { get; set; }         // e.g., "B2:C3"
    }

    class Program
    {
        static void Main()
        {
            try
            {
                // Path to the JSON file that contains shape configurations
                string jsonPath = "shapesConfig.json";

                // Ensure the JSON file exists
                if (!File.Exists(jsonPath))
                {
                    Console.WriteLine($"Configuration file not found: {jsonPath}");
                    return;
                }

                // Read the entire JSON content
                string jsonContent = File.ReadAllText(jsonPath);

                // Deserialize JSON into a list of ShapeConfig objects
                List<ShapeConfig>? shapeConfigs = JsonSerializer.Deserialize<List<ShapeConfig>>(jsonContent);
                if (shapeConfigs == null)
                {
                    Console.WriteLine("Failed to deserialize shape configurations.");
                    return;
                }

                // -------------------------------------------------
                // Create a new workbook (creation rule)
                // -------------------------------------------------
                Workbook workbook = new Workbook();

                // Access the first worksheet where shapes will be recreated
                Worksheet sheet = workbook.Worksheets[0];

                // -------------------------------------------------
                // Recreate each shape based on its configuration
                // -------------------------------------------------
                foreach (var cfg in shapeConfigs)
                {
                    try
                    {
                        // Compute size and offsets. If Width/Height are not provided,
                        // approximate using cell dimensions (rough estimates).
                        int width = cfg.Width > 0
                            ? cfg.Width
                            : (cfg.LowerRightColumn - cfg.UpperLeftColumn + 1) * 64;   // approx 64 px per column
                        int height = cfg.Height > 0
                            ? cfg.Height
                            : (cfg.LowerRightRow - cfg.UpperLeftRow + 1) * 15;        // approx 15 px per row

                        int upperLeftRowOffset = 0;
                        int upperLeftColumnOffset = 0;

                        Shape shape = cfg.Type?.ToLower() switch
                        {
                            "rectangle" => sheet.Shapes.AddRectangle(
                                cfg.UpperLeftRow,
                                cfg.UpperLeftColumn,
                                upperLeftRowOffset,
                                upperLeftColumnOffset,
                                height,
                                width),

                            "line" => sheet.Shapes.AddLine(
                                cfg.UpperLeftRow,
                                cfg.UpperLeftColumn,
                                upperLeftRowOffset,
                                upperLeftColumnOffset,
                                height,
                                width),

                            "oval" => sheet.Shapes.AddOval(
                                cfg.UpperLeftRow,
                                cfg.UpperLeftColumn,
                                upperLeftRowOffset,
                                upperLeftColumnOffset,
                                height,
                                width),

                            // Fallback to a generic rectangle shape
                            _ => sheet.Shapes.AddShape(
                                MsoDrawingType.Rectangle,
                                cfg.UpperLeftRow,
                                cfg.UpperLeftColumn,
                                upperLeftRowOffset,
                                upperLeftColumnOffset,
                                height,
                                width)
                        };

                        // Apply optional size overrides if they were explicitly set
                        if (cfg.Width > 0) shape.Width = cfg.Width;
                        if (cfg.Height > 0) shape.Height = cfg.Height;

                        // Apply common properties
                        if (!string.IsNullOrEmpty(cfg.Name))
                            shape.Name = cfg.Name;

                        if (!string.IsNullOrEmpty(cfg.LinkedCell))
                            shape.SetLinkedCell(cfg.LinkedCell, false, false);

                        if (!string.IsNullOrEmpty(cfg.InputRange))
                            shape.SetInputRange(cfg.InputRange, false, false);
                    }
                    catch (Exception exShape)
                    {
                        Console.WriteLine($"Failed to create shape '{cfg.Name ?? cfg.Type}': {exShape.Message}");
                    }
                }

                // -------------------------------------------------
                // Save the workbook (save rule)
                // -------------------------------------------------
                string outputPath = "ShapesFromJson.xlsx";
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
