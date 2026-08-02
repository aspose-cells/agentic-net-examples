// Title: Aspose.Cells .NET – Add a Rectangle Shape, Retrieve Connection Points, and Export to JSON
// Description: Creates a new workbook, inserts a rectangle shape on the first worksheet, calls GetConnectionPoints() to obtain the shape's anchor coordinates, serializes the X‑Y pairs into a formatted JSON file, and saves both the JSON data and the workbook.
// Keywords: Aspose.Cells shape connection points | GetConnectionPoints C# | export shape coordinates to JSON | add rectangle shape Aspose.Cells | serialize shape points .NET | save workbook with shapes | C# Excel shape API | global
// Common Searches: how to get connection points of a shape using Aspose.Cells | export shape anchor points to JSON in C# | add rectangle shape and retrieve its points Aspose.Cells .NET | save Excel workbook with shape geometry data
// Developer Intent: Extract a shape's connection points after adding it to a worksheet and write the coordinates to a JSON file for downstream processing.
// Use Cases: Feed shape anchor coordinates to a custom diagramming or layout engine. | Create an audit log of geometric data for shapes used in financial reports. | Synchronize Excel shape positions with external systems that require JSON‑formatted geometry.
// AI Prompts: Generate C# code that adds an ellipse shape, reads its connection points with Aspose.Cells, and saves the data to an XML file. | Show how to loop through a shape's connection points and draw connector lines between each successive point on the same worksheet. | Explain how to extend the JSON output to include the shape's ID, type, and bounding box along with each X‑Y coordinate.

using System;
using System.IO;
using System.Text.Json;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a new workbook, inserts a rectangle shape on the first worksheet, calls GetConnectionPoints() to obtain the shape's anchor coordinates, serializes the X‑Y pairs into a formatted JSON file, and saves both the JSON data and the workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet.
            // Parameters: upperLeftRow, upperLeftColumn, lowerRightRow, lowerRightColumn, height (points), width (points)
            Shape shape = worksheet.Shapes.AddRectangle(1, 0, 5, 2, 100, 200);

            // Retrieve the connection points of the shape
            float[][] points = shape.GetConnectionPoints();

            // Prepare a list of anonymous objects for JSON serialization
            List<object> pointList = new List<object>();
            for (int i = 0; i < points.Length; i++)
            {
                pointList.Add(new { X = points[i][0], Y = points[i][1] });
            }

            // Serialize the points list to a formatted JSON string
            string json = JsonSerializer.Serialize(pointList, new JsonSerializerOptions { WriteIndented = true });

            // Write the JSON string to a file
            string jsonPath = "connectionPoints.json";
            File.WriteAllText(jsonPath, json);
            Console.WriteLine($"Connection points saved to '{jsonPath}'.");

            // Save the workbook containing the shape
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
