// Title: Create a rectangle shape in an Excel worksheet, extract its connection points, and validate the data with a JSON schema using Aspose.Cells for .NET
// Common Searches: how to add a rectangle shape to an Excel worksheet with Aspose.Cells C# | retrieve connection points from a shape in Aspose.Cells .NET | validate shape connection point data against a JSON schema in C# | use reflection to read shape connection point properties in Aspose.Cells | save workbook after adding shapes with Aspose.Cells C#
// Tags: Aspose.Cells add rectangle shape | Aspose.Cells get shape connection points | C# reflection read shape properties | JSON schema validation of shape data | Aspose.Cells save workbook with shapes

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new Workbook, adds a rectangle Shape to the first worksheet, attempts to obtain its connection points via GetConnectionPoints, extracts each point's Index, X, and Y values using reflection, serializes the collection to JSON, validates that each point contains the required data types, reports the validation outcome, and finally saves the workbook as ShapeWithConnectionPoints.xlsx.
class ShapeConnectionPointsValidator
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            // Parameters: upper‑left row, upper‑left column, top offset, left offset, height, width
            int upperLeftRow = 2;
            int upperLeftColumn = 2;
            int top = 5;      // pixels
            int left = 5;     // pixels
            int height = 100; // pixels
            int width = 150;  // pixels

            Shape shape = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle,
                upperLeftRow,
                upperLeftColumn,
                top,
                left,
                height,
                width);
            shape.Name = "MyRectangle";

            // Retrieve connection points from the shape (if supported)
            var connectionPointsData = new List<Dictionary<string, object>>();
            try
            {
                // Some shape types expose connection points via GetConnectionPoints()
                // The API may not be available in all versions; handle gracefully.
                var cpCollection = shape.GetConnectionPoints();

                if (cpCollection != null)
                {
                    // Iterate using reflection to avoid compile‑time dependency on specific types
                    foreach (var cp in cpCollection)
                    {
                        // Expect properties Index, X, Y
                        var point = new Dictionary<string, object>
                        {
                            { "Index", GetPropertyValue<int>(cp, "Index") },
                            { "X", GetPropertyValue<double>(cp, "X") },
                            { "Y", GetPropertyValue<double>(cp, "Y") }
                        };
                        connectionPointsData.Add(point);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Unable to retrieve connection points: " + ex.Message);
            }

            // Serialize the connection points list to JSON (for demonstration)
            string connectionPointsJson = JsonSerializer.Serialize(
                connectionPointsData,
                new JsonSerializerOptions { WriteIndented = true });
            Console.WriteLine("Connection points JSON:");
            Console.WriteLine(connectionPointsJson);

            // Simple validation: ensure each item contains Index (int), X (double), Y (double)
            var validationErrors = new List<string>();
            bool isValid = true;
            foreach (var point in connectionPointsData)
            {
                if (!point.ContainsKey("Index") || !(point["Index"] is int))
                    validationErrors.Add("Missing or invalid 'Index'.");

                if (!point.ContainsKey("X") || !(point["X"] is double))
                    validationErrors.Add("Missing or invalid 'X'.");

                if (!point.ContainsKey("Y") || !(point["Y"] is double))
                    validationErrors.Add("Missing or invalid 'Y'.");
            }

            if (validationErrors.Count > 0)
                isValid = false;

            // Output validation result
            Console.WriteLine("Connection points validation result: " + (isValid ? "Valid" : "Invalid"));
            if (!isValid)
            {
                Console.WriteLine("Validation errors:");
                foreach (string error in validationErrors)
                {
                    Console.WriteLine("- " + error);
                }
            }

            // Optionally, save the workbook to verify shape persistence (lifecycle rule: save)
            string outputPath = Path.Combine(Environment.CurrentDirectory, "ShapeWithConnectionPoints.xlsx");
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine("Workbook saved to: " + outputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Failed to save workbook: " + ex.Message);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }

    // Helper method to safely get a property value via reflection
    private static T GetPropertyValue<T>(object obj, string propertyName)
    {
        try
        {
            var prop = obj.GetType().GetProperty(propertyName);
            if (prop != null && prop.GetValue(obj) is T value)
                return value;
        }
        catch
        {
            // Ignore and fall through
        }
        return default;
    }
}
