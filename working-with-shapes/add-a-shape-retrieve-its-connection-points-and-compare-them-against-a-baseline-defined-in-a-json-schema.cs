// Title: Retrieve Aspose.Cells Shape Connection Points (C#) and Validate with a JSON Baseline
// Description: Creates a new workbook, adds a rectangle shape, calls Shape.GetConnectionPoints() to obtain its connection points, maps them to a simple PointModel list, loads expected points from a JSON file that follows a defined schema, compares each point within a tolerance, reports mismatches, and optionally saves the workbook.
// Keywords: Aspose.Cells GetConnectionPoints | C# shape connection points | validate shape geometry JSON | Aspose.Cells rectangle shape example | compare Excel shape points | JSON baseline validation | Excel automation testing Aspose.Cells
// Common Searches: How to get connection points of a shape using Aspose.Cells for .NET | Aspose.Cells compare shape points to JSON file C# | Validate rectangle shape geometry in Excel with Aspose.Cells | Shape.GetConnectionPoints example C# | Unit test shape coordinates Aspose.Cells
// Developer Intent: Add a shape, extract its connection points, and verify them against expected coordinates defined in a JSON schema.
// Use Cases: Automated unit test that confirms generated shapes retain correct connection points across releases. | Quality‑control script that checks diagram shapes in exported Excel files match a design specification stored in JSON. | CI pipeline validation to detect unintended changes in shape geometry after code modifications.
// AI Prompts: Generate a reusable C# method that reads a JSON array of point objects and compares it with the float[][] returned by Shape.GetConnectionPoints(), returning a detailed mismatch report. | Provide example JSON content representing the expected connection points for a default rectangle shape added with Aspose.Cells. | Write robust error‑handling code that logs indices of connection points that differ beyond a tolerance and suggests possible causes.

using System;
using System.IO;
using System.Text.Json;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ShapeConnectionPointsComparison
{
    // Simple model matching the JSON schema for a point
    // Creates a new workbook, adds a rectangle shape, calls Shape.GetConnectionPoints() to obtain its connection points, maps them to a simple PointModel list, loads expected points from a JSON file that follows a defined schema, compares each point within a tolerance, reports mismatches, and optionally saves the workbook.
    public class PointModel
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
            // Parameters: upper left row, upper left column, upper left offset X, upper left offset Y, width, height
            Shape shape = worksheet.Shapes.AddRectangle(1, 0, 0, 100, 200, 0);

            // ---------- Retrieve connection points ----------
            float[][] connectionPoints = shape.GetConnectionPoints();

            // Convert retrieved points to a list of PointModel for easier comparison
            List<PointModel> actualPoints = new List<PointModel>();
            foreach (float[] pt in connectionPoints)
            {
                if (pt.Length >= 2)
                {
                    actualPoints.Add(new PointModel { X = pt[0], Y = pt[1] });
                }
            }

            // ---------- Load baseline points from JSON ----------
            // Expected JSON format: [{ "X": 0.0, "Y": 0.0 }, { "X": 100.0, "Y": 0.0 }, ...]
            string jsonPath = "baseline.json";
            if (!File.Exists(jsonPath))
            {
                Console.WriteLine($"Baseline file not found: {jsonPath}");
                return;
            }

            string jsonContent = File.ReadAllText(jsonPath);
            List<PointModel> baselinePoints = JsonSerializer.Deserialize<List<PointModel>>(jsonContent);

            // ---------- Compare actual points with baseline ----------
            const float tolerance = 0.001f; // allowable difference
            bool allMatch = true;

            if (baselinePoints == null || baselinePoints.Count != actualPoints.Count)
            {
                allMatch = false;
                Console.WriteLine("Point count mismatch between actual and baseline.");
            }
            else
            {
                for (int i = 0; i < baselinePoints.Count; i++)
                {
                    float dx = Math.Abs(baselinePoints[i].X - actualPoints[i].X);
                    float dy = Math.Abs(baselinePoints[i].Y - actualPoints[i].Y);
                    if (dx > tolerance || dy > tolerance)
                    {
                        allMatch = false;
                        Console.WriteLine($"Mismatch at point {i + 1}: Expected ({baselinePoints[i].X}, {baselinePoints[i].Y}) " +
                                          $"but got ({actualPoints[i].X}, {actualPoints[i].Y})");
                    }
                }
            }

            Console.WriteLine(allMatch ? "All connection points match the baseline." : "Connection points do not match the baseline.");

            // ---------- Save the workbook (optional) ----------
            workbook.Save("ShapeConnectionPointsDemo.xlsx");
        }
    }
}
