// Title: Insert a rectangle shape in an Excel workbook with Aspose.Cells, load expected coordinates from a JSON file, and verify the shape’s position in C#
// AI Prompts: Generate C# code that uses Aspose.Cells to add a rectangle shape to a worksheet, reads the expected shape name and coordinates from a JSON configuration, and asserts that the shape's UpperLeftRow and UpperLeftColumn are within a small tolerance of the expected values. | Write a method that deserializes a JSON file containing a shape name and a list of expected points, retrieves the matching shape from an Aspose.Cells workbook, and returns true only if every actual connection point matches the expected point within 0.001 units.
// Common Searches: c# Aspose.Cells verify shape position using JSON configuration file | how to compare Aspose.Cells shape connection points with expected coordinates | read shape expected points from config.json and validate in Aspose.Cells workbook | Aspose.Cells shape UpperLeftRow UpperLeftColumn tolerance check in C# | load shape name and points from JSON and match with workbook shape Aspose.Cells
// Tags: add rectangle shape to worksheet Aspose.Cells C# | deserialize shape config JSON in C# | validate shape coordinates against expected values Aspose.Cells | compare actual shape points with tolerance C# | save workbook after shape verification Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, inserts a rectangle shape named "MyShape", reads the expected shape name and coordinate points from a config.json file, compares the shape's UpperLeftRow and UpperLeftColumn to the expected values within a 0.001 tolerance, reports the match result, and saves the workbook as output.xlsx.
public class ConfigPoint
{
    public double X { get; set; }
    public double Y { get; set; }
}

public class ShapeConfig
{
    public string ShapeName { get; set; } = string.Empty;
    public List<ConfigPoint> ExpectedPoints { get; set; } = new();
}

class Program
{
    static void Main()
    {
        try
        {
            // ---------- Create a new workbook ----------
            var workbook = new Workbook();
            var sheet = workbook.Worksheets[0];

            // ---------- Add a rectangle shape ----------
            // AddShape returns a Shape object in recent Aspose.Cells versions
            Shape shape = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // shape type
                1,    // upper left row
                1,    // upper left column
                0,    // upper left row offset (in points)
                0,    // upper left column offset (in points)
                100,  // height (in points)
                100); // width (in points)

            shape.Name = "MyShape";

            // ---------- Retrieve shape position as a sample point ----------
            var actualPoints = new List<ConfigPoint>
            {
                new ConfigPoint
                {
                    X = shape.UpperLeftRow,
                    Y = shape.UpperLeftColumn
                }
            };

            // ---------- Load expected coordinates from config ----------
            const string configPath = "config.json";
            if (!File.Exists(configPath))
            {
                Console.WriteLine($"Config file not found: {configPath}");
                return;
            }

            string json = File.ReadAllText(configPath);
            ShapeConfig? config = JsonSerializer.Deserialize<ShapeConfig>(json);
            if (config == null)
            {
                Console.WriteLine("Failed to deserialize configuration.");
                return;
            }

            // ---------- Compare actual vs. expected ----------
            bool match = true;

            if (!string.Equals(shape.Name, config.ShapeName, StringComparison.OrdinalIgnoreCase))
            {
                Console.WriteLine($"Shape name mismatch. Workbook shape: {shape.Name}, Config shape: {config.ShapeName}");
                match = false;
            }
            else if (config.ExpectedPoints == null || config.ExpectedPoints.Count != actualPoints.Count)
            {
                Console.WriteLine("Number of points does not match the expected count.");
                match = false;
            }
            else
            {
                const double tolerance = 0.001;
                for (int i = 0; i < actualPoints.Count; i++)
                {
                    double dx = Math.Abs(actualPoints[i].X - config.ExpectedPoints[i].X);
                    double dy = Math.Abs(actualPoints[i].Y - config.ExpectedPoints[i].Y);
                    if (dx > tolerance || dy > tolerance)
                    {
                        Console.WriteLine($"Point {i} mismatch. Actual: ({actualPoints[i].X}, {actualPoints[i].Y}) " +
                                          $"Expected: ({config.ExpectedPoints[i].X}, {config.ExpectedPoints[i].Y})");
                        match = false;
                        break;
                    }
                }
            }

            Console.WriteLine(match
                ? "Points match expected coordinates."
                : "Points do NOT match expected coordinates.");

            // ---------- Save the workbook ----------
            const string outputPath = "output.xlsx";
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
