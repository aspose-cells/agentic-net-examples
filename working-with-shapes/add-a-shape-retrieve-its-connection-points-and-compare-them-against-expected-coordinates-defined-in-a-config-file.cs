// Title: Add a rectangle shape, retrieve its connection points, and validate against a JSON config – Aspose.Cells C# example
// Description: This C# sample creates a workbook, inserts a rectangle shape, calls GetConnectionPoints() to obtain its connection points, loads expected coordinates from a config.json file using System.Text.Json, compares each point within a 0.001 tolerance, reports mismatches, and saves the file as ShapeConnectionPointsDemo.xlsx.
// Keywords: Aspose.Cells shape connection points | C# GetConnectionPoints | validate shape coordinates JSON | rectangle shape geometry tolerance | Aspose.Cells read config file | .NET spreadsheet shape testing
// Common Searches: Aspose.Cells get shape connection points C# | compare shape points with JSON file | validate rectangle geometry Aspose.Cells | load expected shape coordinates from config | C# spreadsheet shape testing tutorial
// Developer Intent: Add a shape, extract its connection points, and confirm they match predefined coordinates stored in a JSON configuration.
// Use Cases: Automated verification of diagram geometry in generated spreadsheets | Compliance checks against design specifications saved in config files | Regression testing for shape placement and connectivity in reporting templates
// AI Prompts: Generate a reusable method that reads expected connection points from a JSON file and compares them with shape.GetConnectionPoints() using a configurable tolerance. | Add robust error handling for missing config files, JSON deserialization errors, and point‑count mismatches when validating shape connection points. | Write NUnit tests that mock Shape.GetConnectionPoints() output and assert that the comparison logic correctly identifies matching and mismatching points.

using System;
using System.IO;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// This C# sample creates a workbook, inserts a rectangle shape, calls GetConnectionPoints() to obtain its connection points, loads expected coordinates from a config.json file using System.Text.Json, compares each point within a 0.001 tolerance, reports mismatches, and saves the file as ShapeConnectionPointsDemo.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a rectangle shape to the worksheet
        // Parameters: upper left row, upper left column, top offset, left offset, height, width, shape type (0 = rectangle)
        Shape shape = worksheet.Shapes.AddRectangle(1, 0, 0, 100, 200, 0);

        // Retrieve the connection points of the shape
        float[][] actualPoints = shape.GetConnectionPoints();

        // Load expected connection points from a JSON config file
        // Expected JSON format:
        // {
        //   "ExpectedPoints": [
        //     [x1, y1],
        //     [x2, y2],
        //     ...
        //   ]
        // }
        string configPath = "config.json";
        if (!File.Exists(configPath))
        {
            Console.WriteLine($"Config file not found: {configPath}");
            return;
        }

        string json = File.ReadAllText(configPath);
        Config config = JsonSerializer.Deserialize<Config>(json);

        // Compare actual points with expected points
        bool allMatch = true;
        const float tolerance = 0.001f; // allowable difference

        if (config.ExpectedPoints.Length != actualPoints.Length)
        {
            allMatch = false;
            Console.WriteLine($"Point count mismatch. Expected {config.ExpectedPoints.Length}, got {actualPoints.Length}");
        }
        else
        {
            for (int i = 0; i < actualPoints.Length; i++)
            {
                float expectedX = config.ExpectedPoints[i][0];
                float expectedY = config.ExpectedPoints[i][1];
                float actualX = actualPoints[i][0];
                float actualY = actualPoints[i][1];

                if (Math.Abs(expectedX - actualX) > tolerance || Math.Abs(expectedY - actualY) > tolerance)
                {
                    allMatch = false;
                    Console.WriteLine($"Point {i + 1} mismatch. Expected ({expectedX}, {expectedY}) but got ({actualX}, {actualY})");
                }
            }
        }

        Console.WriteLine(allMatch ? "All connection points match expected values." : "Some connection points do not match.");

        // Save the workbook
        workbook.Save("ShapeConnectionPointsDemo.xlsx");
    }

    // Helper class for deserializing the config file
    public class Config
    {
        public float[][] ExpectedPoints { get; set; }
    }
}
