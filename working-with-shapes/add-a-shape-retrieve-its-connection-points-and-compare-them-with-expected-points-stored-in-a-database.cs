// Title: Add a Rectangle Shape, Retrieve Its Connection Points, and Validate Them with Aspose.Cells (C#)
// Description: Creates a new Workbook, inserts a rectangle shape on the first worksheet, calls GetConnectionPoints to obtain the shape's normalized connection points, loads a predefined set of expected points, compares both collections with a 0.01 tolerance, reports the outcome, and optionally saves the file. Ideal for regression testing of shape geometry in Excel files using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# shape connection points | GetConnectionPoints | rectangle shape | normalized connection points | shape geometry validation | Excel workbook testing | Aspose.Cells API | regression test
// Common Searches: Aspose.Cells get shape connection points C# | Validate rectangle shape points Aspose.Cells | Retrieve connection points of a shape in .NET | Compare shape connection points with stored data | Aspose.Cells shape geometry testing
// Developer Intent: Add a shape, extract its connection points, and confirm they match a reference set.
// Use Cases: Verify that a newly added rectangle provides the correct eight normalized connection points before further layout processing. | Automate regression tests that compare actual shape points to reference values stored in a database or file. | Log mismatched points to troubleshoot rendering or alignment issues in generated Excel worksheets.
// AI Prompts: Generate a method that reads expected connection points for a shape ID from a SQL database and returns a List<float[]>. | Write an NUnit test that adds a rectangle shape, calls GetConnectionPoints, and asserts the points match expected values within a 0.01 tolerance. | Create code that logs detailed differences between actual and expected connection points when the comparison fails.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapeConnectionPointsDemo
{
    // Creates a new Workbook, inserts a rectangle shape on the first worksheet, calls GetConnectionPoints to obtain the shape's normalized connection points, loads a predefined set of expected points, compares both collections with a 0.01 tolerance, reports the outcome, and optionally saves the file. Ideal for regression testing of shape geometry in Excel files using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            try
            {
                // 1. Create a new workbook and get the first worksheet.
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // 2. Add a rectangle shape to the worksheet.
                // Parameters: upper left row, upper left column, upper left offset X, upper left offset Y, width, height
                Shape shape = worksheet.Shapes.AddRectangle(1, 0, 0, 100, 200, 0);

                // 3. Retrieve the actual connection points from the shape.
                float[][] actualPoints = shape.GetConnectionPoints();

                // 4. Load expected points (shape identifier = 1 for demo).
                int shapeId = 1;
                List<float[]> expectedPoints = LoadExpectedPoints(shapeId);

                // 5. Compare actual points with expected points.
                bool allMatch = ComparePoints(actualPoints, expectedPoints, tolerance: 0.01f);

                Console.WriteLine(allMatch
                    ? "All connection points match the expected values."
                    : "Connection points do NOT match the expected values.");

                // 6. (Optional) Save the workbook to verify the shape was added.
                string outputPath = "ShapeConnectionPoints.xlsx";
                try
                {
                    workbook.Save(outputPath);
                    Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
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

        // Returns hard‑coded expected connection points for demonstration purposes.
        private static List<float[]> LoadExpectedPoints(int shapeId)
        {
            var points = new List<float[]>();

            try
            {
                // For a rectangle shape Aspose.Cells returns 8 connection points:
                // (0,0), (0.5,0), (1,0), (1,0.5), (1,1), (0.5,1), (0,1), (0,0.5)
                // These values are normalized (relative to shape bounds).
                points.Add(new float[] { 0f, 0f });
                points.Add(new float[] { 0.5f, 0f });
                points.Add(new float[] { 1f, 0f });
                points.Add(new float[] { 1f, 0.5f });
                points.Add(new float[] { 1f, 1f });
                points.Add(new float[] { 0.5f, 1f });
                points.Add(new float[] { 0f, 1f });
                points.Add(new float[] { 0f, 0.5f });
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load expected points: {ex.Message}");
            }

            return points;
        }

        private static bool ComparePoints(float[][] actual, List<float[]> expected, float tolerance)
        {
            if (actual == null || expected == null)
                return false;

            if (actual.Length != expected.Count)
                return false;

            for (int i = 0; i < actual.Length; i++)
            {
                float[] act = actual[i];
                float[] exp = expected[i];

                if (act.Length != 2 || exp.Length != 2)
                    return false;

                if (Math.Abs(act[0] - exp[0]) > tolerance || Math.Abs(act[1] - exp[1]) > tolerance)
                    return false;
            }

            return true;
        }
    }
}
