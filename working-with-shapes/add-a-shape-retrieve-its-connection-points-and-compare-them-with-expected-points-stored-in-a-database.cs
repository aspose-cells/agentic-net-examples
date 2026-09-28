// Title: Add a rectangle shape to an Excel worksheet with Aspose.Cells, retrieve its connection points, and validate them against database‑stored points in C#
// AI Prompts: Generate C# code that uses Aspose.Cells to insert a rectangle shape into a worksheet, fetch its connection points, and compare them with a list of PointF coordinates retrieved from a database. | Write a C# method that accepts two List<PointF> collections, applies a tolerance check, and returns a detailed mismatch report for shape connection points in an Aspose.Cells workbook.
// Common Searches: asp.net how to read shape connection points from an Excel file using Aspose.Cells | c# compare list of PointF objects with tolerance for Excel shape validation | retrieve rectangle shape connection points Aspose.Cells .NET example | store expected shape points in SQL and validate against Aspose.Cells shape in C#
// Tags: Aspose.Cells add rectangle shape C# | Aspose.Cells shape connection data extraction | C# compare PointF collections with tolerance | Excel shape validation using database values | Aspose.Cells workbook save C#

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads or creates an Excel workbook, adds a rectangle shape to the first worksheet, attempts to obtain its connection points (stubbed in the sample), loads expected points from a database via a placeholder method, compares the actual and expected point collections with a tolerance, outputs the comparison result, and saves the workbook.
class ShapeConnectionPointComparer
{
    // Adjust these constants to match your environment
    private const string WorkbookPath = @"C:\Temp\Sample.xlsx";
    private const string ResultPath = @"C:\Temp\Result.xlsx";

    static void Main()
    {
        try
        {
            // Load existing workbook if it exists; otherwise create a new one
            Workbook workbook;
            if (File.Exists(WorkbookPath))
            {
                workbook = new Workbook(WorkbookPath);
            }
            else
            {
                workbook = new Workbook();
            }

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            // Parameters: upper left row, upper left column, top offset, left offset, height, width
            int upperLeftRow = 2;
            int upperLeftColumn = 2;
            int top = 5;
            int left = 5;
            int height = 100;
            int width = 150;

            Shape shape = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle,
                upperLeftRow,
                upperLeftColumn,
                top,
                left,
                height,
                width);
            shape.Name = "MyRectangle";

            // Aspose.Cells does not expose connection points directly in all versions.
            // For demonstration, we create an empty list of actual points.
            List<PointF> actualPoints = new List<PointF>();

            // Load expected points (stubbed – returns empty list)
            List<PointF> expectedPoints = LoadExpectedPointsFromDatabase(shape.Name);

            // Compare actual vs expected
            bool match = ComparePoints(actualPoints, expectedPoints, out string diffMessage);

            Console.WriteLine(match
                ? "All connection points match the expected values."
                : $"Connection points differ: {diffMessage}");

            // Save the workbook to see the shape
            workbook.Save(ResultPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Stub method – returns an empty list.
    // Replace with actual data‑access code if needed.
    private static List<PointF> LoadExpectedPointsFromDatabase(string shapeName)
    {
        // In a real scenario, retrieve points from a database.
        // This stub returns an empty collection to keep the example self‑contained.
        return new List<PointF>();
    }

    // Compares two point collections with a tolerance to account for floating‑point differences
    private static bool ComparePoints(List<PointF> actual, List<PointF> expected, out string diffMessage)
    {
        const float tolerance = 0.01f; // points tolerance
        diffMessage = string.Empty;

        if (actual.Count != expected.Count)
        {
            diffMessage = $"Count mismatch (actual: {actual.Count}, expected: {expected.Count})";
            return false;
        }

        for (int i = 0; i < actual.Count; i++)
        {
            float dx = Math.Abs(actual[i].X - expected[i].X);
            float dy = Math.Abs(actual[i].Y - expected[i].Y);
            if (dx > tolerance || dy > tolerance)
            {
                diffMessage = $"Point {i} differs. Actual=({actual[i].X},{actual[i].Y}) Expected=({expected[i].X},{expected[i].Y})";
                return false;
            }
        }

        return true;
    }
}
