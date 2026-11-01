// Title: Create a rectangle shape in Aspose.Cells, retrieve its connection points via reflection, and validate them against a JSON baseline (C#)
// AI Prompts: Write C# code that adds a rectangle shape to an Aspose.Cells worksheet, uses reflection to invoke the hidden GetConnectionPoints method, and compares each returned point with coordinates defined in a baseline JSON file. | Generate a .NET console program that loads a JSON file containing expected connection‑point indices and coordinates, extracts the actual points from an Aspose.Cells shape through reflection, and reports any mismatches within a small tolerance.
// Common Searches: how to get shape connection points in Aspose.Cells when GetConnectionPoints is not public | C# compare Aspose.Cells shape connection points with values from a JSON file | validate rectangle shape geometry in Aspose.Cells using a baseline JSON schema | using reflection to access hidden shape methods in Aspose.Cells .NET | Aspose.Cells retrieve connection points for a rectangle shape and check tolerance
// Tags: Aspose.Cells shape connection points reflection | C# rectangle shape validation with JSON | compare shape geometry baseline Aspose.Cells | retrieve shape connection points .NET | JSON baseline verification for Aspose.Cells shapes

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The program creates a new workbook, adds a rectangle shape, loads a baseline JSON file with expected connection‑point indices and coordinates, uses reflection to call GetConnectionPoints on the shape, compares each point to the baseline within a tolerance, reports mismatches, and saves the workbook as an .xlsx file.
class ShapeConnectionPointComparer
{
    static void Main()
    {
        try
        {
            // ---------- Create a new workbook ----------
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // ---------- Add a rectangle shape ----------
            // Row 2, Column 2 (0‑based), size 100x50 points
            Shape shape = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle,
                2, 2,
                0, 0,
                100, 50);

            // ---------- Load baseline JSON ----------
            string jsonPath = "baseline.json";
            if (!File.Exists(jsonPath))
            {
                Console.WriteLine($"Baseline file not found: {jsonPath}");
                return;
            }

            string jsonContent = File.ReadAllText(jsonPath);
            using JsonDocument doc = JsonDocument.Parse(jsonContent);
            JsonElement root = doc.RootElement;

            if (!root.TryGetProperty("connectionPoints", out JsonElement pointsArray) ||
                pointsArray.ValueKind != JsonValueKind.Array)
            {
                Console.WriteLine("Invalid baseline JSON format: missing 'connectionPoints' array.");
                return;
            }

            // Build a dictionary for quick lookup by index
            var baselineDict = new Dictionary<int, (double X, double Y)>();
            foreach (JsonElement point in pointsArray.EnumerateArray())
            {
                int idx = point.GetProperty("index").GetInt32();
                double x = point.GetProperty("x").GetDouble();
                double y = point.GetProperty("y").GetDouble();
                baselineDict[idx] = (x, y);
            }

            // ---------- Retrieve connection points via reflection ----------
            // Use reflection to stay compatible with versions that may not expose the API directly.
            var getConnPointsMethod = shape.GetType().GetMethod("GetConnectionPoints");
            if (getConnPointsMethod == null)
            {
                Console.WriteLine("The current Aspose.Cells version does not support shape connection points.");
            }
            else
            {
                var connPointsObj = getConnPointsMethod.Invoke(shape, null) as System.Collections.IEnumerable;
                if (connPointsObj == null)
                {
                    Console.WriteLine("Failed to retrieve connection points.");
                }
                else
                {
                    const double tolerance = 0.0001;
                    bool allMatch = true;

                    foreach (var cpObj in connPointsObj)
                    {
                        // Use dynamic to access Index, X, Y without compile‑time type dependency
                        dynamic cp = cpObj;
                        int idx = cp.Index;
                        double shapeX = cp.X;
                        double shapeY = cp.Y;

                        if (baselineDict.TryGetValue(idx, out var baseline))
                        {
                            bool xMatch = Math.Abs(shapeX - baseline.X) <= tolerance;
                            bool yMatch = Math.Abs(shapeY - baseline.Y) <= tolerance;

                            if (!xMatch || !yMatch)
                            {
                                allMatch = false;
                                Console.WriteLine($"Mismatch at index {idx}: Shape({shapeX}, {shapeY}) vs Baseline({baseline.X}, {baseline.Y})");
                            }
                        }
                        else
                        {
                            allMatch = false;
                            Console.WriteLine($"Baseline does not contain connection point with index {idx}");
                        }
                    }

                    if (allMatch)
                        Console.WriteLine("All connection points match the baseline.");
                }
            }

            // ---------- Save the workbook ----------
            string outputPath = "ShapeWithConnectionPoints.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
