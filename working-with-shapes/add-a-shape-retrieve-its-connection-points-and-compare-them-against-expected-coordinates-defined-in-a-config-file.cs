// Title: Add a rectangle shape with Aspose.Cells, retrieve its connection points, and validate them against a JSON configuration file (C#)
// AI Prompts: Write C# code that inserts a rectangle shape using Aspose.Cells, reads expected connection point data from a JSON file, and verifies each point with Shape.GetConnectionPoints(). | Demonstrate how to change the shape type, position, or size while preserving the JSON‑driven connection‑point validation logic. | Create a reusable method that logs any mismatched connection points, returns a validation flag, and saves the workbook to a given path.
// Common Searches: how to get connection points of a shape using Aspose.Cells C# | compare Excel shape coordinates with values from a JSON config in .NET | validate shape geometry in Aspose.Cells workbook against expected points | read expected shape connection points from a config file and verify in C# | Aspose.Cells Shape.GetConnectionPoints example with JSON validation
// Tags: Aspose.Cells retrieve shape connection points | validate shape geometry with JSON Aspose.Cells | add rectangle shape and compare points C# | Shape.GetConnectionPoints verification | JSON driven shape coordinate validation Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ShapeConnectionPointsDemo
{
    // Classes to map the configuration file
    // The sample creates a new workbook, adds a rectangle shape to the first worksheet, loads expected connection point coordinates from a JSON configuration file, retrieves the shape's actual connection points via GetConnectionPoints(), compares each point's row, column, offsetX, and offsetY with the expected values, logs any mismatches, and saves the workbook.
    public class Config
    {
        public List<ExpectedPoint> ExpectedPoints { get; set; } = new();
    }

    public class ExpectedPoint
    {
        public int Row { get; set; }
        public int Column { get; set; }
        public int OffsetX { get; set; }
        public int OffsetY { get; set; }
    }

    class Program
    {
        static void Main()
        {
            try
            {
                // ----- Load expected coordinates from config file -----
                const string configPath = "config.json";
                if (!File.Exists(configPath))
                {
                    Console.WriteLine($"Config file not found: {configPath}");
                    return;
                }

                string json = File.ReadAllText(configPath);
                Config? config = JsonSerializer.Deserialize<Config>(json);
                if (config == null || config.ExpectedPoints == null)
                {
                    Console.WriteLine("Failed to deserialize configuration.");
                    return;
                }

                // ----- Create a new workbook and add a shape -----
                Workbook workbook = new Workbook();                     // Create a new workbook
                Worksheet sheet = workbook.Worksheets[0];              // Get the first worksheet

                // Add a rectangle shape (you can change the type as needed)
                // Parameters: type, upperRow, upperColumn, upperRowOffset, lowerRow, lowerColumn, lowerRowOffset
                Shape shape = sheet.Shapes.AddShape(
                    MsoDrawingType.Rectangle, // Shape type
                    2, 1, 0,                  // Upper left cell (row 2, column 1) with no offset
                    5, 4, 0);                 // Lower right cell (row 5, column 4) with no offset

                shape.Width = 120;   // Optional: set explicit width
                shape.Height = 80;   // Optional: set explicit height

                // ----- Retrieve connection points of the shape -----
                // The method returns an array of connection point objects.
                var actualPoints = shape.GetConnectionPoints();

                // ----- Compare actual points with expected points from config -----
                bool allMatch = true;

                if (actualPoints.Length != config.ExpectedPoints.Count)
                {
                    allMatch = false;
                    Console.WriteLine($"Count mismatch: actual={actualPoints.Length}, expected={config.ExpectedPoints.Count}");
                }
                else
                {
                    for (int i = 0; i < actualPoints.Length; i++)
                    {
                        // Use dynamic to avoid compile‑time dependency on the exact type name.
                        dynamic actual = actualPoints[i];
                        ExpectedPoint expected = config.ExpectedPoints[i];

                        if (actual.Row != expected.Row ||
                            actual.Column != expected.Column ||
                            actual.OffsetX != expected.OffsetX ||
                            actual.OffsetY != expected.OffsetY)
                        {
                            allMatch = false;
                            Console.WriteLine($"Point {i} mismatch:");
                            Console.WriteLine($"  Actual   -> Row:{actual.Row}, Column:{actual.Column}, OffsetX:{actual.OffsetX}, OffsetY:{actual.OffsetY}");
                            Console.WriteLine($"  Expected -> Row:{expected.Row}, Column:{expected.Column}, OffsetX:{expected.OffsetX}, OffsetY:{expected.OffsetY}");
                        }
                    }
                }

                Console.WriteLine(allMatch
                    ? "All connection points match the expected coordinates."
                    : "Some connection points do not match the expected coordinates.");

                // ----- Save the workbook -----
                const string outputPath = "ShapeWithConnectionPoints.xlsx";
                try
                {
                    workbook.Save(outputPath);
                    Console.WriteLine($"Workbook saved to {outputPath}");
                }
                catch (Exception saveEx)
                {
                    Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
