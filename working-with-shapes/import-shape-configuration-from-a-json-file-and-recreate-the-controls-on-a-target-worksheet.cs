// Title: Import JSON Shape Configurations and Recreate Excel Shapes with Aspose.Cells for .NET (C#)
// Description: A C# example that reads a JSON file of shape definitions (type, name, row/column position, offsets, size, optional text), deserializes them into ShapeConfig objects, creates matching rectangle, line or oval shapes on the first worksheet using Aspose.Cells drawing APIs, applies custom properties, handles missing or empty files gracefully, and saves the result as an Excel workbook.
// Keywords: Aspose.Cells | C# | .NET | JSON | shape import | Excel shapes | rectangle | line | oval | drawing API | workbook save | shape configuration | deserialize JSON | add shape programmatically
// Common Searches: Aspose.Cells read shape settings from JSON | Create Excel shapes from external config C# | Deserialize shape list and add to worksheet Aspose.Cells | Import rectangle line oval shapes via JSON | How to generate shapes in Excel using Aspose.Cells and JSON
// Developer Intent: Load shape parameters from a JSON file and generate the same shapes on an Excel worksheet with Aspose.Cells.
// Use Cases: Generate a workbook that reproduces rectangles, lines, and ovals defined in a JSON configuration. | Produce an empty workbook when the JSON file is missing or contains no shapes, avoiding runtime errors. | Assign custom names and text to each shape for later identification, editing, or automation.
// AI Prompts: Write C# code that reads a JSON array of shape configurations and adds the corresponding shapes to an Aspose.Cells worksheet, including error handling for unsupported types. | Provide a JSON schema that matches the ShapeConfig class used for importing shapes into Aspose.Cells. | Explain how to extend the sample to support additional shape types such as Triangle or Picture with Aspose.Cells.

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ShapeImportExample
{
    // Represents a shape configuration read from JSON.
    // A C# example that reads a JSON file of shape definitions (type, name, row/column position, offsets, size, optional text), deserializes them into ShapeConfig objects, creates matching rectangle, line or oval shapes on the first worksheet using Aspose.Cells drawing APIs, applies custom properties, handles missing or empty files gracefully, and saves the result as an Excel workbook.
    public class ShapeConfig
    {
        public string Type { get; set; }               // e.g., "Rectangle", "Line", "Oval"
        public string Name { get; set; }               // Shape name
        public int UpperLeftRow { get; set; }          // Starting row (0‑based)
        public int UpperLeftColumn { get; set; }       // Starting column (0‑based)
        public int UpperLeftRowOffset { get; set; }    // Pixel offset from the upper‑left row
        public int UpperLeftColumnOffset { get; set; } // Pixel offset from the upper‑left column
        public int Height { get; set; }                // Height in pixels
        public int Width { get; set; }                 // Width in pixels
        public string Text { get; set; }               // Optional text for the shape
    }

    class Program
    {
        static void Main()
        {
            try
            {
                // Path to the JSON file that contains an array of shape configurations.
                const string jsonPath = "shapesConfig.json";

                // Verify that the JSON file exists before attempting to read it.
                if (!File.Exists(jsonPath))
                {
                    Console.WriteLine($"Configuration file '{jsonPath}' not found. No shapes will be created.");
                    // Proceed with an empty list so the workbook is still generated.
                    ProcessShapes(new List<ShapeConfig>());
                    return;
                }

                // Read and deserialize the JSON content.
                string jsonContent = File.ReadAllText(jsonPath);
                List<ShapeConfig> shapeConfigs = JsonSerializer.Deserialize<List<ShapeConfig>>(jsonContent) 
                                                ?? new List<ShapeConfig>();

                ProcessShapes(shapeConfigs);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Creates shapes based on the provided configurations and saves the workbook.
        private static void ProcessShapes(List<ShapeConfig> shapeConfigs)
        {
            // Create a new workbook.
            Workbook workbook = new Workbook();

            // Access the first worksheet where shapes will be recreated.
            Worksheet sheet = workbook.Worksheets[0];

            // Iterate over each shape configuration and create the corresponding shape.
            foreach (var cfg in shapeConfigs)
            {
                if (cfg == null) continue;

                Shape shape = null;
                switch (cfg.Type?.Trim().ToLower())
                {
                    case "rectangle":
                        shape = sheet.Shapes.AddRectangle(
                            cfg.UpperLeftRow,
                            cfg.UpperLeftColumn,
                            cfg.UpperLeftRowOffset,
                            cfg.UpperLeftColumnOffset,
                            cfg.Height,
                            cfg.Width);
                        break;
                    case "line":
                        shape = sheet.Shapes.AddLine(
                            cfg.UpperLeftRow,
                            cfg.UpperLeftColumn,
                            cfg.UpperLeftRowOffset,
                            cfg.UpperLeftColumnOffset,
                            cfg.Height,
                            cfg.Width);
                        break;
                    case "oval":
                        shape = sheet.Shapes.AddOval(
                            cfg.UpperLeftRow,
                            cfg.UpperLeftColumn,
                            cfg.UpperLeftRowOffset,
                            cfg.UpperLeftColumnOffset,
                            cfg.Height,
                            cfg.Width);
                        break;
                    default:
                        Console.WriteLine($"Unsupported shape type: {cfg.Type}");
                        continue;
                }

                // Apply optional properties.
                if (!string.IsNullOrEmpty(cfg.Name))
                    shape.Name = cfg.Name;

                if (!string.IsNullOrEmpty(cfg.Text))
                    shape.Text = cfg.Text;
            }

            // Save the workbook to an Excel file.
            const string outputPath = "RecreatedShapes.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
    }
}
