using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ShapeConnectionPointsDemo
{
    // Simple model for expected points loaded from config
    public class Config
    {
        public List<PointData> Points { get; set; }
    }

    public class PointData
    {
        public float X { get; set; }
        public float Y { get; set; }
    }

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
            float[][] connectionPoints = shape.GetConnectionPoints();

            // Load expected points from a JSON config file
            // Example config.json content:
            // { "Points": [ { "X": 0, "Y": 0 }, { "X": 100, "Y": 0 }, { "X": 100, "Y": 200 }, { "X": 0, "Y": 200 } ] }
            string configPath = "config.json";
            if (!File.Exists(configPath))
            {
                Console.WriteLine($"Config file not found: {configPath}");
                return;
            }

            Config config = JsonSerializer.Deserialize<Config>(File.ReadAllText(configPath));

            // Compare retrieved points with expected points
            const float tolerance = 0.01f; // tolerance for floating point comparison
            bool allMatch = true;

            if (config.Points == null || config.Points.Count != connectionPoints.Length)
            {
                Console.WriteLine("Mismatch in number of connection points.");
                allMatch = false;
            }
            else
            {
                for (int i = 0; i < connectionPoints.Length; i++)
                {
                    float actualX = connectionPoints[i][0];
                    float actualY = connectionPoints[i][1];
                    float expectedX = config.Points[i].X;
                    float expectedY = config.Points[i].Y;

                    bool matchX = Math.Abs(actualX - expectedX) <= tolerance;
                    bool matchY = Math.Abs(actualY - expectedY) <= tolerance;

                    if (!matchX || !matchY)
                    {
                        Console.WriteLine($"Point {i + 1} mismatch. Expected ({expectedX}, {expectedY}), Actual ({actualX}, {actualY})");
                        allMatch = false;
                    }
                }
            }

            if (allMatch)
            {
                Console.WriteLine("All connection points match the expected coordinates.");
            }

            // Save the workbook
            workbook.Save("ShapeConnectionPointsDemo.xlsx");
        }
    }
}