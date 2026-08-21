// Title: C# – Export Aspose.Cells ShapePath Geometry to JSON
// Description: Demonstrates how to create a freeform shape with ShapePath, access its CustomGeometry, extract segment types and point coordinates, map them to DTO objects, and serialize the result to an indented JSON string using System.Text.Json in Aspose.Cells.
// Keywords: Aspose.Cells ShapePath JSON | C# export shape geometry | CustomGeometry to JSON | freeform shape coordinates | System.Text.Json serialization | Aspose.Cells shape segment extraction | Excel shape path serialization
// Common Searches: Aspose.Cells convert ShapePath to JSON | C# extract freeform shape coordinates from Excel | How to serialize custom geometry in Aspose.Cells | Export shape segment data as JSON using Aspose.Cells | ShapePath JSON example for Aspose.Cells
// Developer Intent: Retrieve the segments and points of a ShapePath from a freeform shape in an Excel workbook and serialize the data to JSON.
// Use Cases: Send shape geometry to a web front‑end for client‑side diagram rendering. | Store shape path coordinates in a database for version control or analytics. | Integrate shape data with third‑party vector‑graphics services via JSON payloads.
// AI Prompts: Write C# code that reads all ShapePath objects from an Aspose.Cells worksheet and outputs their segment types and coordinates to a JSON file using System.Text.Json. | Create a method that converts a collection of ShapeSegmentPath objects into DTOs ready for JSON serialization. | Explain how to handle multiple freeform shapes in a workbook and combine their ShapePath data into a single JSON structure.

using System;
using System.Collections.Generic;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapePathToJson
{
    // DTO classes for JSON serialization
    // Demonstrates how to create a freeform shape with ShapePath, access its CustomGeometry, extract segment types and point coordinates, map them to DTO objects, and serialize the result to an indented JSON string using System.Text.Json in Aspose.Cells.
    public class ShapePathDto
    {
        public List<SegmentDto> Segments { get; set; } = new List<SegmentDto>();
    }

    public class SegmentDto
    {
        public string Type { get; set; }
        public List<PointDto> Points { get; set; } = new List<PointDto>();
    }

    public class PointDto
    {
        public float X { get; set; }
        public float Y { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Build a simple rectangular freeform shape using ShapePath
            ShapePath rectPath = new ShapePath();
            rectPath.MoveTo(10, 10);
            rectPath.LineTo(200, 10);
            rectPath.LineTo(200, 100);
            rectPath.LineTo(10, 100);
            rectPath.Close();

            // Add the freeform shape to the worksheet
            Shape freeform = worksheet.Shapes.AddFreeform(0, 0, 0, 0, 300, 200, new ShapePath[] { rectPath });

            // Cast the shape's geometry to CustomGeometry to access its paths
            CustomGeometry geometry = freeform.Geometry as CustomGeometry;
            if (geometry == null)
            {
                Console.WriteLine("The shape does not contain custom geometry.");
                return;
            }

            // Assume we are interested in the first path (the rectangle we created)
            ShapePath shapePath = geometry.Paths[0];

            // Prepare DTO for JSON
            ShapePathDto dto = new ShapePathDto();

            // Iterate over each segment in the path
            foreach (ShapeSegmentPath segment in shapePath.PathSegementList)
            {
                SegmentDto segDto = new SegmentDto
                {
                    Type = segment.Type.ToString()
                };

                // Collect points of the segment
                foreach (ShapePathPoint pt in segment.Points)
                {
                    segDto.Points.Add(new PointDto { X = pt.X, Y = pt.Y });
                }

                dto.Segments.Add(segDto);
            }

            // Serialize the DTO to JSON (indented for readability)
            string json = JsonSerializer.Serialize(dto, new JsonSerializerOptions { WriteIndented = true });

            // Output the JSON string
            Console.WriteLine(json);

            // Save the workbook (optional, demonstrates lifecycle rule compliance)
            workbook.Save("ShapePathDemo.xlsx");
        }
    }
}
