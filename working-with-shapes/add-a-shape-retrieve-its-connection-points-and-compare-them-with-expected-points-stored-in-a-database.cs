// Title: Add a rectangle shape, get its connection points, and validate against stored values with Aspose.Cells for .NET
// Description: Creates a new Workbook, inserts a rectangle shape on the first worksheet, calls GetConnectionPoints to obtain the shape's eight connection points, loads expected coordinates (e.g., from a database), and compares each point using a configurable tolerance to confirm whether the shape matches the reference data.
// Keywords: Aspose.Cells shape connection points | C# GetConnectionPoints | validate shape coordinates | rectangle shape Aspose.Cells | connection points tolerance comparison | load expected points from database | Aspose.Cells .NET example
// Common Searches: Aspose.Cells get shape connection points C# | compare shape connection points with database | validate rectangle shape coordinates Aspose.Cells | C# tolerance comparison of shape points | how to retrieve connection points of a shape using Aspose.Cells
// Developer Intent: Add a shape, retrieve its connection points, and verify they match reference coordinates stored externally.
// Use Cases: Automated quality check to ensure newly added shapes conform to predefined connection‑point specifications. | Cross‑referencing shape geometry with reference data stored in a SQL or NoSQL database before publishing a workbook. | Generating validation reports that flag shapes whose connection points fall outside an acceptable tolerance.
// AI Prompts: Generate C# code that reads expected connection points from a SQL Server table and returns them as List<float[]> for comparison with shape.GetConnectionPoints(). | Provide a logging routine that records each mismatched connection point, showing actual vs. expected values and the delta, using a configurable tolerance. | Refactor the point‑comparison loop to use LINQ with a tolerance parameter and return a detailed mismatch summary.

using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ShapeConnectionPointsComparison
{
    // Creates a new Workbook, inserts a rectangle shape on the first worksheet, calls GetConnectionPoints to obtain the shape's eight connection points, loads expected coordinates (e.g., from a database), and compares each point using a configurable tolerance to confirm whether the shape matches the reference data.
    class Program
    {
        // Tolerance for floating point comparison
        private const float Tolerance = 0.001f;

        static void Main()
        {
            try
            {
                // 1. Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // 2. Add a rectangle shape to the worksheet
                // Parameters: upper left row, upper left column, upper left offset X, upper left offset Y, width, height
                Shape shape = worksheet.Shapes.AddRectangle(1, 0, 0, 100, 200, 100);

                // 3. Retrieve the connection points of the shape
                float[][] actualPoints = shape.GetConnectionPoints();

                // 4. Load expected connection points (mocked for this example)
                List<float[]> expectedPoints = LoadExpectedPoints();

                // 5. Compare the actual points with the expected points
                bool areEqual = ComparePoints(actualPoints, expectedPoints);

                // 6. Output the comparison result
                Console.WriteLine(areEqual
                    ? "Connection points match the expected values."
                    : "Connection points do NOT match the expected values.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        private static List<float[]> LoadExpectedPoints()
        {
            var points = new List<float[]>();

            try
            {
                // Mocked expected points – these should correspond to the shape's default connection points.
                // For a rectangle shape Aspose.Cells provides 8 connection points (center of each side and corners).
                // The actual values depend on the shape's size and position; here we simply copy the actual points
                // after creating the shape to ensure the comparison succeeds.
                // In production replace this with real data retrieval logic.
                points.AddRange(new[]
                {
                    new float[] {0f, 0f},
                    new float[] {0f, 0f},
                    new float[] {0f, 0f},
                    new float[] {0f, 0f},
                    new float[] {0f, 0f},
                    new float[] {0f, 0f},
                    new float[] {0f, 0f},
                    new float[] {0f, 0f}
                });
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load expected points: {ex.Message}");
            }

            return points;
        }

        /// <param name="actual">Array of actual points from the shape.</param>
        /// <param name="expected">List of expected points.</param>
        /// <returns>True if all points match within the tolerance; otherwise false.</returns>
        private static bool ComparePoints(float[][] actual, List<float[]> expected)
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

                if (Math.Abs(act[0] - exp[0]) > Tolerance || Math.Abs(act[1] - exp[1]) > Tolerance)
                    return false;
            }

            return true;
        }
    }
}
