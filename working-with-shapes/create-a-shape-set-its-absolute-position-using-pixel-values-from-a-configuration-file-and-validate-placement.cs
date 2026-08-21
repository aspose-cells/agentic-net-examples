// Title: C# – Add a Rectangle Shape in Aspose.Cells, Set Absolute Pixel Position from Config File, and Verify Placement
// Description: Shows how to read X and Y pixel coordinates from a simple key‑value configuration file, create a workbook with Aspose.Cells, insert a rectangle shape, assign its X and Y properties in pixels, optionally set Placement to MoveAndSize, validate the coordinates, and save the result as an XLSX file.
// Keywords: Aspose.Cells C# shape position pixels | set shape X Y Aspose.Cells | read shape coordinates config file | validate shape placement Aspose.Cells | add rectangle shape workbook | shape placement type MoveAndSize | pixel‑based shape positioning | Excel shape absolute coordinates | load configuration file C#
// Common Searches: Aspose.Cells set shape position in pixels | C# read shape coordinates from text file | place shape at exact pixel offset in Excel using Aspose.Cells | validate shape X Y values after setting in Aspose.Cells | MoveAndSize placement for shapes Aspose.Cells | load key=value config in C# for Aspose.Cells
// Developer Intent: The developer wants to programmatically position a shape at a specific pixel offset defined in an external configuration file, confirm that the placement matches the expected values, and generate a correctly formatted Excel workbook.
// Use Cases: Insert a company logo at a precise pixel location defined by a config file to maintain branding consistency across generated reports. | Position a dynamic watermark based on user‑provided X/Y coordinates before exporting Excel files. | Automate layout verification by comparing expected and actual shape coordinates in generated spreadsheets. | Create template‑driven Excel documents where shape positions are controlled via external settings for flexible design updates.
// AI Prompts: Generate C# code that reads X and Y pixel positions from a JSON configuration file and sets a rectangle shape's X and Y properties in Aspose.Cells, including robust error handling. | Show how to fall back to default coordinates when the configuration file is missing or contains invalid numbers while positioning a shape. | Provide a method that logs a warning if the shape's actual X or Y differs from the expected values after assignment. | Create a C# unit test that verifies shape placement matches configuration values using Aspose.Cells.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapePlacementDemo
{
    // Shows how to read X and Y pixel coordinates from a simple key‑value configuration file, create a workbook with Aspose.Cells, insert a rectangle shape, assign its X and Y properties in pixels, optionally set Placement to MoveAndSize, validate the coordinates, and save the result as an XLSX file.
    class Program
    {
        static void Main()
        {
            // ---------- Load configuration ----------
            // Expected format (each line): key=value
            // Example:
            // X=150
            // Y=200
            string configPath = "shapeConfig.txt";
            var config = LoadConfig(configPath);

            // Parse pixel values, fallback to defaults if missing or invalid
            int posX = config.ContainsKey("X") && int.TryParse(config["X"], out var x) ? x : 0;
            int posY = config.ContainsKey("Y") && int.TryParse(config["Y"], out var y) ? y : 0;

            // ---------- Create workbook ----------
            Workbook workbook = new Workbook();                     // create new workbook
            Worksheet worksheet = workbook.Worksheets[0];          // get first worksheet

            // ---------- Add a rectangle shape ----------
            // Parameters: upper left column, upper left row, upper left offset X, upper left offset Y, width, height
            // We'll place it initially at (0,0) and then set absolute pixel offsets.
            Shape shape = worksheet.Shapes.AddRectangle(0, 0, 0, 0, 100, 50);

            // ---------- Set absolute position using pixel values ----------
            shape.X = posX;    // horizontal offset from worksheet left border (pixels)
            shape.Y = posY;    // vertical offset from worksheet top border (pixels)

            // Optional: define how the shape moves with cells
            shape.Placement = PlacementType.MoveAndSize;

            // ---------- Validate placement ----------
            bool isXValid = shape.X == posX;
            bool isYValid = shape.Y == posY;

            Console.WriteLine($"Shape X set to {shape.X} (expected {posX}) - Valid: {isXValid}");
            Console.WriteLine($"Shape Y set to {shape.Y} (expected {posY}) - Valid: {isYValid}");

            // ---------- Save workbook ----------
            string outputPath = "ShapePlacementDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }

        // Helper method to read simple key=value configuration file
        private static Dictionary<string, string> LoadConfig(string path)
        {
            var dict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (!File.Exists(path))
            {
                Console.WriteLine($"Configuration file '{path}' not found. Using defaults.");
                return dict;
            }

            foreach (var line in File.ReadAllLines(path))
            {
                var trimmed = line.Trim();
                if (string.IsNullOrEmpty(trimmed) || trimmed.StartsWith("#"))
                    continue; // skip empty lines and comments

                var parts = trimmed.Split(new[] { '=' }, 2);
                if (parts.Length == 2)
                {
                    dict[parts[0].Trim()] = parts[1].Trim();
                }
            }
            return dict;
        }
    }
}
