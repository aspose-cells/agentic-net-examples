using System;
using System.IO;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class ShapeConnectionPointsToJson
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a rectangle shape to the worksheet
        // Parameters: upperLeftRow, top, upperLeftColumn, left, height, width, shapeId
        Shape shape = worksheet.Shapes.AddRectangle(1, 0, 0, 100, 200, 0);

        // Retrieve the connection points of the shape
        float[][] points = shape.GetConnectionPoints();

        // Transform the points into a serializable structure
        var pointList = new System.Collections.Generic.List<object>();
        for (int i = 0; i < points.Length; i++)
        {
            pointList.Add(new { X = points[i][0], Y = points[i][1] });
        }

        // Serialize the list to JSON
        string json = JsonSerializer.Serialize(pointList, new JsonSerializerOptions { WriteIndented = true });

        // Write JSON to a file
        string jsonPath = "ShapeConnectionPoints.json";
        File.WriteAllText(jsonPath, json);
        Console.WriteLine($"Connection points saved to {jsonPath}");

        // Optionally save the workbook (not required for JSON export)
        workbook.Save("ShapeWithConnectionPoints.xlsx");
    }
}