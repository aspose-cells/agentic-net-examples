// Title: Add a rectangle shape to an Excel worksheet with Aspose.Cells, retrieve its connection points, and validate them against a JSON file in C#
// Common Searches: how to get connection points of a shape added with Aspose.Cells in C# | compare shape connection points to values from a JSON config using Aspose.Cells | validate rectangle shape coordinates with tolerance in Aspose.Cells C# example | read expected point coordinates from JSON and check against Aspose.Cells shape points | Aspose.Cells GetConnectionPoints returns float array how to convert to PointF
// Tags: Aspose.Cells add rectangle shape | Aspose.Cells GetConnectionPoints usage | C# compare shape points with JSON data | Excel shape connection points validation | JSON deserialization to PointF list .NET

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, adds a rectangle shape, extracts its connection points via GetConnectionPoints, reads expected points from a JSON configuration file, compares each actual point to the expected coordinates within a small tolerance, reports mismatches, and saves the workbook as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            // Parameters: drawing type, upper left row, upper left column,
            // lower right row, lower right column, width, height
            Shape shape = worksheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // shape type
                2,    // upper left row
                2,    // upper left column
                5,    // lower right row
                5,    // lower right column
                100,  // width (in points)
                100); // height (in points)

            // Retrieve the shape's connection points
            // GetConnectionPoints returns float[][] where each inner array is {X, Y}
            float[][] rawPoints = shape.GetConnectionPoints();

            // Convert to PointF[]
            PointF[] actualPoints = new PointF[rawPoints.Length];
            for (int i = 0; i < rawPoints.Length; i++)
            {
                if (rawPoints[i] != null && rawPoints[i].Length >= 2)
                {
                    actualPoints[i] = new PointF(rawPoints[i][0], rawPoints[i][1]);
                }
                else
                {
                    actualPoints[i] = new PointF(0, 0);
                }
            }

            // Load expected connection points from a JSON config file
            // Expected file format: [{ "X": 100.0, "Y": 200.0 }, ...]
            string configPath = "expectedPoints.json";
            if (!File.Exists(configPath))
            {
                Console.WriteLine($"Config file not found: {configPath}");
                return;
            }

            List<PointF>? expectedPoints = null;
            try
            {
                string json = File.ReadAllText(configPath);
                expectedPoints = JsonSerializer.Deserialize<List<PointF>>(json);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error reading or deserializing config file: {ex.Message}");
                return;
            }

            if (expectedPoints == null)
            {
                Console.WriteLine("Failed to deserialize expected points.");
                return;
            }

            // Compare actual points with expected points
            const float tolerance = 0.01f; // allowable deviation
            bool allMatch = true;

            int compareCount = Math.Min(actualPoints.Length, expectedPoints.Count);
            for (int i = 0; i < compareCount; i++)
            {
                if (Math.Abs(actualPoints[i].X - expectedPoints[i].X) > tolerance ||
                    Math.Abs(actualPoints[i].Y - expectedPoints[i].Y) > tolerance)
                {
                    Console.WriteLine($"Mismatch at index {i}: " +
                                      $"Actual ({actualPoints[i].X}, {actualPoints[i].Y}) vs " +
                                      $"Expected ({expectedPoints[i].X}, {expectedPoints[i].Y})");
                    allMatch = false;
                }
            }

            if (allMatch && actualPoints.Length == expectedPoints.Count)
            {
                Console.WriteLine("All connection points match the expected coordinates.");
            }
            else if (allMatch)
            {
                Console.WriteLine("Point count mismatch between actual and expected data.");
            }

            // Save the workbook to a file
            string outputPath = "output.xlsx";
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
