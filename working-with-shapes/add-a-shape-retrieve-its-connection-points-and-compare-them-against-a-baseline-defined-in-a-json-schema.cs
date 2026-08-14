// Title: Validate Aspose.Cells Rectangle Shape Connection Points Against a JSON Baseline (C#)
// Description: Creates a workbook, adds a rectangle shape, extracts its connection points with GetConnectionPoints(), loads expected coordinates from a baseline.json file, compares each pair within a 0.001 tolerance, reports mismatches, and optionally saves the file.
// Keywords: Aspose.Cells | C# | shape connection points | GetConnectionPoints | JSON baseline | rectangle shape | coordinate comparison | tolerance check | regression testing | spreadsheet template validation
// Common Searches: Aspose.Cells get connection points C# | compare shape points with JSON Aspose | validate rectangle shape geometry .NET | shape connection points tolerance comparison | load baseline points from JSON C#
// Developer Intent: Add a rectangle shape, retrieve its connection points, and confirm they match a predefined set of coordinates stored in a JSON file.
// Use Cases: Automated regression tests that verify shape geometry in generated spreadsheets. | Ensuring custom workbook templates adhere to design specifications by checking shape connection points. | Generating documentation that cross‑references extracted shape coordinates with a design baseline.
// AI Prompts: Write C# code using Aspose.Cells to insert a rectangle shape, read its connection points, and compare them to a JSON array with a 0.001 tolerance. | Explain how to deserialize a JSON array of point coordinates in .NET and use it to validate shape connection points obtained via Aspose.Cells. | Provide a method that logs any mismatched connection points between an Aspose.Cells shape and a baseline JSON, returning a boolean indicating overall success.

using System;
using System.IO;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ShapeConnectionPointsComparison
{
    // Creates a workbook, adds a rectangle shape, extracts its connection points with GetConnectionPoints(), loads expected coordinates from a baseline.json file, compares each pair within a 0.001 tolerance, reports mismatches, and optionally saves the file.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            // Parameters: upper left row, upper left column, upper left offset X, upper left offset Y, width, height, rotation angle
            Shape shape = worksheet.Shapes.AddRectangle(1, 0, 0, 100, 200, 0);

            // Retrieve the connection points of the shape
            float[][] actualPoints = shape.GetConnectionPoints();

            // Load baseline connection points from a JSON file (baseline.json)
            // Expected JSON format: { "points": [ [x1, y1], [x2, y2], ... ] }
            string jsonPath = "baseline.json";
            if (!File.Exists(jsonPath))
            {
                Console.WriteLine($"Baseline file '{jsonPath}' not found.");
                return;
            }

            string jsonContent = File.ReadAllText(jsonPath);
            JsonDocument doc = JsonDocument.Parse(jsonContent);
            JsonElement root = doc.RootElement;

            // Parse baseline points
            JsonElement pointsElement = root.GetProperty("points");
            float[][] baselinePoints = new float[pointsElement.GetArrayLength()][];
            int idx = 0;
            foreach (JsonElement point in pointsElement.EnumerateArray())
            {
                float x = point[0].GetSingle();
                float y = point[1].GetSingle();
                baselinePoints[idx++] = new float[] { x, y };
            }

            // Compare actual points with baseline points
            const float tolerance = 0.001f; // allowable difference
            bool allMatch = true;

            if (actualPoints.Length != baselinePoints.Length)
            {
                Console.WriteLine($"Point count mismatch: actual={actualPoints.Length}, baseline={baselinePoints.Length}");
                allMatch = false;
            }
            else
            {
                for (int i = 0; i < actualPoints.Length; i++)
                {
                    float actualX = actualPoints[i][0];
                    float actualY = actualPoints[i][1];
                    float expectedX = baselinePoints[i][0];
                    float expectedY = baselinePoints[i][1];

                    bool xMatch = Math.Abs(actualX - expectedX) <= tolerance;
                    bool yMatch = Math.Abs(actualY - expectedY) <= tolerance;

                    if (!xMatch || !yMatch)
                    {
                        Console.WriteLine($"Point {i + 1} mismatch: actual=({actualX}, {actualY}) vs expected=({expectedX}, {expectedY})");
                        allMatch = false;
                    }
                }
            }

            if (allMatch)
            {
                Console.WriteLine("All connection points match the baseline.");
            }

            // Save the workbook (optional, demonstrates lifecycle usage)
            workbook.Save("ShapeConnectionPoints.xlsx");
        }
    }
}
