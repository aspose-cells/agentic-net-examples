// Title: Validate Rectangle Shape Connection Points with Aspose.Cells C# and a JSON Config
// Description: Creates a workbook, adds a rectangle shape, extracts its connection points via Shape.GetConnectionPoints, loads expected X/Y values from a JSON file, compares them within a 0.001 tolerance, logs mismatches, and saves the file.
// Keywords: Aspose.Cells | C# | Shape.GetConnectionPoints | rectangle shape | JSON configuration | coordinate validation | tolerance comparison | spreadsheet testing | workbook save | geometry verification
// Common Searches: Aspose.Cells retrieve shape connection points C# | compare shape coordinates with JSON Aspose.Cells | validate rectangle shape geometry .NET | Shape.GetConnectionPoints example | automated shape layout test Aspose.Cells
// Developer Intent: Ensure that the rectangle shape added to a spreadsheet has connection points that exactly match the coordinates defined in an external JSON configuration.
// Use Cases: Automated regression test to confirm shape geometry remains consistent after code changes. | Quality‑control check that programmatically generated diagrams follow design specifications. | Debugging tool that reports any deviation in shape connection points after applying transformations.
// AI Prompts: Generate C# code using Aspose.Cells to add a rectangle shape, retrieve its connection points, and compare them to expected X/Y values from a JSON file with a 0.001 tolerance. | Explain the structure of the float[][] array returned by Shape.GetConnectionPoints and how to map each sub‑array to X and Y coordinates for different shape types. | Create a reusable method that logs mismatched connection points, throws an exception on tolerance breaches, and returns a boolean indicating overall match.

using System;
using System.IO;
using System.Text.Json;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ShapeConnectionPointsComparison
{
    // Simple model to hold expected point coordinates from config
    // Creates a workbook, adds a rectangle shape, extracts its connection points via Shape.GetConnectionPoints, loads expected X/Y values from a JSON file, compares them within a 0.001 tolerance, logs mismatches, and saves the file.
    public class ExpectedPoint
    {
        public float X { get; set; }
        public float Y { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // ---------- Create a new workbook ----------
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // ---------- Add a rectangle shape ----------
            // Parameters: upper left row, upper left column, top offset, left offset, height, width, shape index (0 for default)
            Shape shape = worksheet.Shapes.AddRectangle(1, 0, 0, 100, 200, 0);

            // ---------- Retrieve connection points from the shape ----------
            float[][] actualPoints = shape.GetConnectionPoints();

            // ---------- Load expected points from configuration file ----------
            // Expected config format (JSON):
            // [
            //   { "X": 0.0, "Y": 0.0 },
            //   { "X": 100.0, "Y": 0.0 },
            //   ...
            // ]
            string configPath = "config.json";
            if (!File.Exists(configPath))
            {
                Console.WriteLine($"Config file not found: {configPath}");
                return;
            }

            List<ExpectedPoint> expectedPoints;
            try
            {
                string json = File.ReadAllText(configPath);
                expectedPoints = JsonSerializer.Deserialize<List<ExpectedPoint>>(json);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to read or parse config file: {ex.Message}");
                return;
            }

            // ---------- Compare actual points with expected points ----------
            const float tolerance = 0.001f; // allowable difference
            bool allMatch = true;

            if (actualPoints.Length != expectedPoints.Count)
            {
                Console.WriteLine($"Point count mismatch. Actual: {actualPoints.Length}, Expected: {expectedPoints.Count}");
                allMatch = false;
            }

            int compareCount = Math.Min(actualPoints.Length, expectedPoints.Count);
            for (int i = 0; i < compareCount; i++)
            {
                float actualX = actualPoints[i][0];
                float actualY = actualPoints[i][1];
                float expectedX = expectedPoints[i].X;
                float expectedY = expectedPoints[i].Y;

                bool xMatch = Math.Abs(actualX - expectedX) <= tolerance;
                bool yMatch = Math.Abs(actualY - expectedY) <= tolerance;

                if (!xMatch || !yMatch)
                {
                    Console.WriteLine($"Point {i + 1} mismatch:");
                    Console.WriteLine($"  Expected: X={expectedX}, Y={expectedY}");
                    Console.WriteLine($"  Actual:   X={actualX}, Y={actualY}");
                    allMatch = false;
                }
            }

            if (allMatch)
            {
                Console.WriteLine("All connection points match the expected configuration.");
            }

            // ---------- Save the workbook ----------
            string outputPath = "ShapeConnectionPointsDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
    }
}
