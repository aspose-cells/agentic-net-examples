// Title: Insert a rectangle shape with Aspose.Cells, read its connection points via reflection, and export the points to a JSON file (C#)
// AI Prompts: Generate C# code that adds a rectangle shape to a worksheet using Aspose.Cells, accesses the shape's ConnectionPoints collection through reflection, and writes each point's index and X/Y coordinates to a formatted JSON file. | Create a C# program that creates a new workbook, inserts a shape, uses reflection to retrieve any hidden ConnectionPoints, and serializes the collected points list to an indented JSON document while also saving the workbook. | Provide a C# snippet that adds a rectangle shape to an Excel file, extracts the shape's connection point data (Index, X, Y) via reflection, and outputs the data to a pretty‑printed JSON file.
// Common Searches: Aspose.Cells C# get shape connection points using reflection | export Excel shape connection points to JSON with Aspose.Cells | how to retrieve rectangle shape coordinates from Aspose.Cells workbook | C# save shape connection point data as JSON file | access hidden ConnectionPoints property of Aspose.Cells shape
// Tags: Aspose.Cells insert rectangle shape | Aspose.Cells retrieve connection points via reflection | C# serialize shape points to JSON | export Excel shape coordinates JSON | reflection access hidden shape properties Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;

// The example creates a new workbook, inserts a rectangle shape on the first worksheet, uses reflection to locate and iterate the shape's ConnectionPoints collection (if present), gathers each point's Index, X, and Y values into a list, serializes the list to an indented JSON file, and saves both the JSON file and the workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            // Parameters: shape type, upper left row, upper left column, top offset, left offset, height, width
            Shape shape = worksheet.Shapes.AddShape(MsoDrawingType.Rectangle, 2, 2, 0, 0, 100, 50);

            // Retrieve the shape's connection points (if supported)
            var connectionPoints = new List<object>();

            // Aspose.Cells may not expose ConnectionPoints in some versions.
            // Guard against missing API by checking via reflection.
            var cpProperty = shape.GetType().GetProperty("ConnectionPoints");
            if (cpProperty != null)
            {
                var cpCollection = cpProperty.GetValue(shape) as System.Collections.IEnumerable;
                if (cpCollection != null)
                {
                    foreach (var cpObj in cpCollection)
                    {
                        var indexProp = cpObj.GetType().GetProperty("Index");
                        var xProp = cpObj.GetType().GetProperty("X");
                        var yProp = cpObj.GetType().GetProperty("Y");

                        if (indexProp != null && xProp != null && yProp != null)
                        {
                            connectionPoints.Add(new
                            {
                                Index = indexProp.GetValue(cpObj),
                                X = xProp.GetValue(cpObj),
                                Y = yProp.GetValue(cpObj)
                            });
                        }
                    }
                }
            }

            // Serialize the connection points list to JSON
            string json = JsonSerializer.Serialize(connectionPoints, new JsonSerializerOptions { WriteIndented = true });

            // Write the JSON to a file
            string jsonFilePath = "ShapeConnectionPoints.json";
            File.WriteAllText(jsonFilePath, json);

            // Optionally save the workbook to verify the shape was added
            string workbookPath = "ShapeWorkbook.xlsx";
            workbook.Save(workbookPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
