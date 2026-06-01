using System;
using System.IO;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class ShapeConnectionPointsValidator
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            // Parameters: shape type, upper left row, upper left column, row offset, column offset, height, width
            Shape shape = worksheet.Shapes.AddShape(MsoDrawingType.Rectangle, 1, 0, 0, 0, 100, 200);

            // Retrieve the connection points of the shape
            float[][] actualPoints = shape.GetConnectionPoints();

            // Load baseline connection points from a JSON schema file (baseline.json)
            const string baselinePath = "baseline.json";
            if (!File.Exists(baselinePath))
            {
                Console.WriteLine($"Baseline file not found: {baselinePath}");
                return;
            }

            string json = File.ReadAllText(baselinePath);
            using JsonDocument doc = JsonDocument.Parse(json);
            JsonElement root = doc.RootElement;
            JsonElement pointsElement = root.GetProperty("points");

            // Parse baseline points into a float[][] array
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
            bool match = true;

            if (actualPoints.Length != baselinePoints.Length)
            {
                Console.WriteLine($"Point count mismatch: actual={actualPoints.Length}, baseline={baselinePoints.Length}");
                match = false;
            }
            else
            {
                for (int i = 0; i < actualPoints.Length; i++)
                {
                    float actualX = actualPoints[i][0];
                    float actualY = actualPoints[i][1];
                    float expectedX = baselinePoints[i][0];
                    float expectedY = baselinePoints[i][1];

                    if (Math.Abs(actualX - expectedX) > tolerance || Math.Abs(actualY - expectedY) > tolerance)
                    {
                        Console.WriteLine($"Point {i + 1} differs. Actual: ({actualX}, {actualY}) Expected: ({expectedX}, {expectedY})");
                        match = false;
                    }
                }
            }

            Console.WriteLine(match ? "All connection points match the baseline." : "Connection points do not match the baseline.");

            // Save the workbook (lifecycle: save)
            const string outputPath = "ShapeConnectionPoints.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}