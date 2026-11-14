// Title: Create a rectangle shape in an Excel worksheet, read its connection points, and export them to a JSON file using Aspose.Cells for .NET
// Common Searches: how to get connection points of a shape in Aspose.Cells C# | export shape connection point coordinates to JSON with Aspose.Cells | add rectangle shape to Excel workbook and retrieve its connection points using .NET | serialize Aspose.Cells shape connection points to a JSON file in C# | Aspose.Cells get shape connection points list and save as JSON
// Tags: Aspose.Cells add rectangle shape | Aspose.Cells retrieve shape connection points | Aspose.Cells export shape data to JSON | C# serialize shape connection points | Excel workbook shape connection points extraction

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, adds a rectangle shape to the first worksheet, obtains the shape's connection points, builds a collection of each point's index and X/Y coordinates, serializes this collection to formatted JSON, writes the JSON to a file, and saves the workbook containing the shape.
class ShapeConnectionPointsExport
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            // Parameters: shape type, upper left row, upper left column, upper left offset X, upper left offset Y, width, height
            Shape shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 2, 2, 0, 0, 150, 80);

            // Retrieve the collection of connection points from the shape
            // GetConnectionPoints() returns a collection that can be enumerated; use var to avoid version‑specific type issues
            var points = shape.GetConnectionPoints();

            // Prepare a list to hold the connection point data
            var pointList = new List<object>();

            // Iterate through each connection point and collect its index and coordinates
            foreach (var pt in points)
            {
                // Use dynamic to access properties safely across different library versions
                dynamic point = pt;
                pointList.Add(new
                {
                    Index = point.Index,
                    X = point.X,
                    Y = point.Y
                });
            }

            // Serialize the list to JSON
            string json = JsonSerializer.Serialize(pointList, new JsonSerializerOptions { WriteIndented = true });

            // Define the output JSON file path
            string jsonPath = "ShapeConnectionPoints.json";

            // Write the JSON string to the file
            File.WriteAllText(jsonPath, json);
            Console.WriteLine($"Connection points exported to '{jsonPath}'.");

            // Optionally, save the workbook to see the shape in Excel (not required for JSON export)
            string workbookPath = "ShapeWithConnectionPoints.xlsx";
            workbook.Save(workbookPath);
            Console.WriteLine($"Workbook saved to '{workbookPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
