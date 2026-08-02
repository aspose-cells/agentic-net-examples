// Title: Aspose.Cells C# – Set Shape Absolute Position from a Config File and Verify Placement
// Description: Loads X/Y pixel values from a simple text config, adds a zero‑size rectangle shape to the first worksheet, assigns shape.X and shape.Y, prints actual vs. expected coordinates, and saves the workbook.
// Keywords: Aspose.Cells shape position | C# set shape X Y pixels | load shape coordinates config file | absolute pixel placement Aspose.Cells | validate shape location | Excel rectangle shape positioning | shape placement property | read config file C# | pixel based shape layout | Aspose.Cells US developers | Aspose.Cells Europe examples
// Common Searches: How to position a shape in Aspose.Cells using pixel values | Read X and Y coordinates from a text file for an Aspose.Cells shape | Set absolute X and Y for a rectangle shape in C# | Validate shape coordinates after assigning them in Aspose.Cells | Configure shape Placement property after setting pixel position
// Developer Intent: Add a rectangle shape, place it at exact pixel offsets read from a configuration file, and confirm the coordinates are applied correctly.
// Use Cases: Insert a company logo at a fixed pixel offset on generated reports for consistent branding. | Align watermarks across multiple workbooks by pulling coordinates from a shared config file. | Automated QA that checks shape positions before distributing Excel files to avoid layout errors.
// AI Prompts: Generate C# code that reads X and Y values from a JSON config and sets the shape's X/Y properties in Aspose.Cells. | Show how to change the Shape.Placement enum after setting absolute pixel coordinates to control behavior on sheet resize. | Provide error‑handling logic for missing, non‑numeric, or out‑of‑range coordinate values when positioning a shape. | Create a reusable method that loads shape coordinates from any key‑value file and applies them to multiple shapes in a workbook.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Loads X/Y pixel values from a simple text config, adds a zero‑size rectangle shape to the first worksheet, assigns shape.X and shape.Y, prints actual vs. expected coordinates, and saves the workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Load pixel positions from a simple config file (e.g., X=150, Y=80)
            var config = LoadConfig("shapeConfig.txt");
            int x = int.Parse(config["X"]);
            int y = int.Parse(config["Y"]);

            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a rectangle shape with zero size; we'll set its absolute position later
            Shape shape = sheet.Shapes.AddRectangle(0, 0, 0, 0, 0, 0);

            // Set absolute position using pixel values from the config
            shape.X = x; // horizontal offset from worksheet left border
            shape.Y = y; // vertical offset from worksheet top border

            // Validate that the shape was positioned as expected
            Console.WriteLine($"Shape X position: {shape.X} (expected {x})");
            Console.WriteLine($"Shape Y position: {shape.Y} (expected {y})");
            Console.WriteLine($"Shape Placement: {shape.Placement}");

            // Save the workbook
            string outputPath = "ShapePositionDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Simple parser for a key=value text file
    static Dictionary<string, string> LoadConfig(string path)
    {
        var dict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        if (!File.Exists(path))
        {
            Console.WriteLine($"Config file '{path}' not found. Using default values X=150, Y=80.");
            dict["X"] = "150";
            dict["Y"] = "80";
            return dict;
        }

        foreach (var line in File.ReadAllLines(path))
        {
            var trimmed = line.Trim();
            if (string.IsNullOrEmpty(trimmed) || trimmed.StartsWith("#"))
                continue;

            var parts = trimmed.Split(new[] { '=' }, 2);
            if (parts.Length == 2)
                dict[parts[0].Trim()] = parts[1].Trim();
        }

        // Ensure required keys exist; provide defaults if missing
        if (!dict.ContainsKey("X")) dict["X"] = "150";
        if (!dict.ContainsKey("Y")) dict["Y"] = "80";

        return dict;
    }
}
