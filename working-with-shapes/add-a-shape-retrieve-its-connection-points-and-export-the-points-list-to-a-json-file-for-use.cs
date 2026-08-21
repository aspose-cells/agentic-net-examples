// Title: C# – Export Shape Connection Points to JSON with Aspose.Cells
// Description: Creates a workbook, adds a rectangle shape, retrieves its connection points via GetConnectionPoints(), converts the coordinates to a serializable list, writes the list to a formatted JSON file, and optionally saves the workbook containing the shape.
// Keywords: Aspose.Cells C# shape connection points | GetConnectionPoints Aspose.Cells | export shape coordinates to JSON | Aspose.Cells serialize shape geometry | C# write JSON file | Aspose.Cells rectangle shape example | shape anchor points JSON | Aspose.Cells .NET API | shape geometry export | connection points list
// Common Searches: Aspose.Cells GetConnectionPoints example C# | How to export shape coordinates to JSON in .NET | C# Aspose.Cells shape connection points | Save shape anchor points as JSON file | Export all shape connection points Aspose.Cells
// Developer Intent: Extract a shape's connection points and store them in a JSON file using Aspose.Cells for .NET.
// Use Cases: Generate a JSON map of shape anchor points for custom diagram layout engines. | Share shape geometry with external reporting tools that consume JSON data. | Version‑control shape connection data for later reconstruction in another workbook.
// AI Prompts: Show how to include each shape's name and ID alongside its connection points in the exported JSON. | Provide code that reads the generated ShapeConnectionPoints.json and recreates the shape's connection points in a new workbook. | Explain how to iterate over all shapes on a worksheet and combine their connection points into a single JSON array.

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a workbook, adds a rectangle shape, retrieves its connection points via GetConnectionPoints(), converts the coordinates to a serializable list, writes the list to a formatted JSON file, and optionally saves the workbook containing the shape.
class ExportShapeConnectionPoints
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            // Parameters: upperLeftRow, upperLeftColumn, upperLeftRowOffset, upperLeftColumnOffset, height, width
            Shape shape = worksheet.Shapes.AddRectangle(1, 1, 0, 0, 100, 200);

            // Retrieve the connection points of the shape
            float[][] points = shape.GetConnectionPoints();

            // Transform the points into a serializable structure
            var pointList = new List<object>();
            for (int i = 0; i < points.Length; i++)
            {
                pointList.Add(new { X = points[i][0], Y = points[i][1] });
            }

            // Serialize the points list to JSON
            string json = JsonSerializer.Serialize(pointList, new JsonSerializerOptions { WriteIndented = true });

            // Write the JSON to a file
            string jsonPath = "ShapeConnectionPoints.json";
            File.WriteAllText(jsonPath, json);
            Console.WriteLine($"Connection points exported to {jsonPath}");

            // Save the workbook (optional, to keep the shape in the file)
            string workbookPath = "ShapeWithConnectionPoints.xlsx";
            workbook.Save(workbookPath);
            Console.WriteLine($"Workbook saved to {workbookPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
