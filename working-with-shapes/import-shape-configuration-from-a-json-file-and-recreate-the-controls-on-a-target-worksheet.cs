// Title: Read shape definitions from a JSON file and add Rectangle, Oval, and TextBox shapes to an Excel worksheet using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads a JSON array of shape objects and uses Aspose.Cells to create matching Rectangle, Oval, or TextBox shapes on a specified worksheet. | Show how to map JSON shape type strings to Aspose.Cells MsoDrawingType values and assign position, size, and optional text to each created shape. | Demonstrate robust error handling for missing JSON files, malformed JSON, and unsupported shape types while saving the workbook with Aspose.Cells.
// Common Searches: Aspose.Cells C# import shapes from JSON and draw them on an Excel worksheet | How to deserialize shape configuration and add rectangles and textboxes with Aspose.Cells | C# code to read shape properties from a file and create Excel shapes using Aspose.Cells API | Add oval shapes to Excel using Aspose.Cells based on external JSON data
// Tags: json shape deserialization Aspose.Cells C# | create rectangle oval textbox Aspose.Cells | add shapes to Excel worksheet Aspose.Cells | handle missing json file Aspose.Cells | map json type to MsoDrawingType Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ShapeImportExample
{
    // Represents the configuration for a shape read from JSON.
    // The example reads a JSON file containing a list of shape configurations (type, position, size, and optional text), loads or creates an Excel workbook, and iterates through the configurations to add corresponding Rectangle, Oval, or TextBox shapes to the first worksheet using Aspose.Cells. It sets the text property when provided, includes error handling for file and JSON issues, and saves the result as a new workbook.
    public class ShapeConfig
    {
        public string? Type { get; set; }          // e.g., "Rectangle", "Oval", "TextBox"
        public double Left { get; set; }           // Position from the left edge (in points)
        public double Top { get; set; }            // Position from the top edge (in points)
        public double Width { get; set; }          // Width of the shape (in points)
        public double Height { get; set; }         // Height of the shape (in points)
        public string? Text { get; set; }          // Optional text for the shape
    }

    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Path to the JSON file containing shape configurations.
                const string jsonPath = "shapes.json";

                // Load and deserialize the JSON file into a list of ShapeConfig objects.
                List<ShapeConfig> shapeConfigs = LoadShapeConfigurations(jsonPath);
                if (shapeConfigs == null)
                {
                    Console.WriteLine("No shape configurations were loaded.");
                    return;
                }

                // Load the target workbook (or create a new one if it does not exist).
                const string workbookPath = "TargetWorkbook.xlsx";
                Workbook workbook = File.Exists(workbookPath)
                    ? new Workbook(workbookPath)   // Load existing workbook
                    : new Workbook();              // Create a new workbook

                // Choose the worksheet where shapes will be added.
                Worksheet sheet = workbook.Worksheets[0]; // Using the first worksheet

                // Recreate each shape on the worksheet based on the configuration.
                foreach (var config in shapeConfigs)
                {
                    if (string.IsNullOrWhiteSpace(config.Type))
                        continue; // Skip entries without a type

                    Shape shape = null;

                    // Aspose.Cells expects integer row/column indices; using 0 for both.
                    int row = 0;
                    int column = 0;
                    int top = (int)config.Top;
                    int left = (int)config.Left;
                    int height = (int)config.Height;
                    int width = (int)config.Width;

                    switch (config.Type.Trim().ToLower())
                    {
                        case "rectangle":
                            shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, row, column, top, left, height, width);
                            break;
                        case "oval":
                        case "ellipse":
                            shape = sheet.Shapes.AddShape(MsoDrawingType.Oval, row, column, top, left, height, width);
                            break;
                        case "textbox":
                            shape = sheet.Shapes.AddTextBox(row, column, top, left, height, width);
                            break;
                        // Add more shape types as needed.
                        default:
                            // Unknown type – skip.
                            continue;
                    }

                    // Apply common properties.
                    if (!string.IsNullOrEmpty(config.Text) && shape != null)
                    {
                        shape.Text = config.Text;
                    }

                    // Additional formatting can be applied here (e.g., fill color, line style).
                    // Example: shape.FillColor = System.Drawing.Color.LightBlue;
                }

                // Save the workbook with the newly added shapes.
                workbook.Save("ResultWorkbook.xlsx");
                Console.WriteLine("Workbook saved as ResultWorkbook.xlsx");
            }
            catch (FileNotFoundException ex)
            {
                Console.WriteLine($"File error: {ex.Message}");
            }
            catch (JsonException ex)
            {
                Console.WriteLine($"JSON parsing error: {ex.Message}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }

        // Helper method to load shape configurations from a JSON file.
        private static List<ShapeConfig> LoadShapeConfigurations(string jsonFilePath)
        {
            if (!File.Exists(jsonFilePath))
                throw new FileNotFoundException($"JSON configuration file not found: {jsonFilePath}");

            string jsonContent = File.ReadAllText(jsonFilePath);
            var options = new JsonSerializerOptions
            {
                PropertyNameCaseInsensitive = true
            };
            return JsonSerializer.Deserialize<List<ShapeConfig>>(jsonContent, options) ?? new List<ShapeConfig>();
        }
    }
}
