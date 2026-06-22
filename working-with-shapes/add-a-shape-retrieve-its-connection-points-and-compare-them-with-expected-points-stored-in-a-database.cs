using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using System.IO;

namespace ShapeConnectionPointsDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a rectangle shape to the worksheet
                // Parameters: upper left row, upper left column, upper left offset X, upper left offset Y, width, height, shape type (0 = rectangle)
                Shape shape = worksheet.Shapes.AddRectangle(1, 0, 0, 100, 200, 0);

                // Retrieve the connection points of the shape
                float[][] connectionPoints = shape.GetConnectionPoints();

                // Load expected connection points.
                // In a real scenario this could come from a database; here we use a fallback list.
                List<float[]> expectedPoints = LoadExpectedPoints();

                // Compare the retrieved connection points with the expected points
                bool allMatch = true;
                int count = Math.Min(connectionPoints.Length, expectedPoints.Count);

                for (int i = 0; i < count; i++)
                {
                    float actualX = connectionPoints[i][0];
                    float actualY = connectionPoints[i][1];
                    float expectedX = expectedPoints[i][0];
                    float expectedY = expectedPoints[i][1];

                    // Allow a small tolerance for floating point comparison
                    const float tolerance = 0.001f;
                    bool match = Math.Abs(actualX - expectedX) <= tolerance &&
                                 Math.Abs(actualY - expectedY) <= tolerance;

                    Console.WriteLine($"Point {i + 1}: Actual=({actualX}, {actualY}) Expected=({expectedX}, {expectedY}) Match={match}");

                    if (!match)
                    {
                        allMatch = false;
                    }
                }

                // Report any discrepancy in count of points
                if (connectionPoints.Length != expectedPoints.Count)
                {
                    Console.WriteLine($"Warning: Number of connection points ({connectionPoints.Length}) does not match expected count ({expectedPoints.Count}).");
                    allMatch = false;
                }

                Console.WriteLine(allMatch ? "All connection points match expected values." : "There are mismatches in connection points.");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors to prevent the program from crashing
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        /// <summary>
        /// Loads expected connection points.
        /// This method first tries to read a CSV file named \"ExpectedPoints.csv\" located in the executable folder.
        /// If the file does not exist, it returns a default list with placeholder values.
        /// </summary>
        private static List<float[]> LoadExpectedPoints()
        {
            var points = new List<float[]>();
            string filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ExpectedPoints.csv");

            if (File.Exists(filePath))
            {
                try
                {
                    foreach (var line in File.ReadAllLines(filePath))
                    {
                        // Expected CSV format: X,Y
                        var parts = line.Split(',');
                        if (parts.Length >= 2 &&
                            float.TryParse(parts[0], out float x) &&
                            float.TryParse(parts[1], out float y))
                        {
                            points.Add(new float[] { x, y });
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to read expected points from file: {ex.Message}");
                }
            }
            else
            {
                // Fallback: use the default connection points that Aspose.Cells provides for a rectangle shape
                // These values are typical but may vary; they serve as a safe placeholder.
                points.Add(new float[] { 0f, 0f });
                points.Add(new float[] { 0f, 1f });
                points.Add(new float[] { 1f, 1f });
                points.Add(new float[] { 1f, 0f });
            }

            return points;
        }
    }
}