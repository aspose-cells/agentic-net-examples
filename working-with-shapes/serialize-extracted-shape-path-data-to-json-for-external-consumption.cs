// Title: C# – Serialize Aspose.Cells ShapePath Geometry to JSON
// Description: This example creates a workbook, adds a rectangular free‑form shape, extracts its CustomGeometry, iterates through each ShapePath and its segments, captures segment types and point coordinates, builds a serializable object with width, height and segment data, and writes the formatted JSON to the console, a file, and optionally saves the workbook.
// Keywords: Aspose.Cells | C# | ShapePath | CustomGeometry | JSON serialization | freeform shape | export shape coordinates | Excel shape data | GitHub example | Aspose.Cells API
// Common Searches: export Aspose.Cells shape path to JSON | C# get freeform shape coordinates Aspose.Cells | serialize custom geometry Aspose.Cells | convert ShapePath to JSON in .NET | extract shape segment points from Excel workbook
// Developer Intent: Generate a JSON file that contains the dimensions, segment types, and point coordinates of a free‑form shape’s custom geometry created with Aspose.Cells.
// Use Cases: Send shape geometry to a web service that renders diagrams from coordinate data. | Store shape outlines in a database for versioning or reconstruction in other workbooks. | Compare shape outlines across multiple Excel files by serializing them to JSON for automated validation.
// AI Prompts: Write C# code that reads the produced shapePathData.json and recreates the original free‑form shape in a new Aspose.Cells workbook. | Provide a method to deserialize the JSON back into Aspose.Cells ShapePath and ShapeSegment objects while preserving segment types and coordinates. | Explain how to extend the serialization to include fill color, line style, and other visual properties for each shape.

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// This example creates a workbook, adds a rectangular free‑form shape, extracts its CustomGeometry, iterates through each ShapePath and its segments, captures segment types and point coordinates, builds a serializable object with width, height and segment data, and writes the formatted JSON to the console, a file, and optionally saves the workbook.
class SerializeShapePathToJson
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Build a simple rectangular shape path
        ShapePath shapePath = new ShapePath();
        shapePath.MoveTo(10, 10);
        shapePath.LineTo(200, 10);
        shapePath.LineTo(200, 100);
        shapePath.LineTo(10, 100);
        shapePath.Close();

        // Add the shape path as a freeform shape to the worksheet
        Shape shape = worksheet.Shapes.AddFreeform(0, 0, 0, 0, 300, 200, new ShapePath[] { shapePath });

        // Cast the shape geometry to CustomGeometry to access its paths
        CustomGeometry geometry = shape.Geometry as CustomGeometry;
        if (geometry == null)
        {
            Console.WriteLine("The shape does not contain custom geometry.");
            return;
        }

        // Prepare a serializable structure for all paths and their segments
        var pathsData = new List<object>();

        foreach (ShapePath sp in geometry.Paths)
        {
            var segmentList = new List<object>();

            foreach (ShapeSegmentPath segment in sp.PathSegementList)
            {
                var pointList = new List<object>();
                foreach (ShapePathPoint pt in segment.Points)
                {
                    pointList.Add(new { X = pt.X, Y = pt.Y });
                }

                segmentList.Add(new
                {
                    Type = segment.Type.ToString(),
                    Points = pointList
                });
            }

            pathsData.Add(new
            {
                WidthPixel = sp.WidthPixel,
                HeightPixel = sp.HeightPixel,
                Segments = segmentList
            });
        }

        // Serialize the structure to JSON with indentation
        string json = JsonSerializer.Serialize(pathsData, new JsonSerializerOptions { WriteIndented = true });

        // Output JSON to console and write it to a file
        Console.WriteLine(json);
        File.WriteAllText("shapePathData.json", json);

        // Save the workbook (optional, demonstrates lifecycle usage)
        workbook.Save("ShapePathDemo.xlsx");
    }
}
