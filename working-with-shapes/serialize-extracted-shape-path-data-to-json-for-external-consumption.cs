using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsDemo
{
    class SerializeShapePathDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Build a simple rectangular shape path
                ShapePath path = new ShapePath();
                path.MoveTo(10, 10);
                path.LineTo(200, 10);
                path.LineTo(200, 100);
                path.LineTo(10, 100);
                path.Close();

                // Add the freeform shape that uses the created path
                sheet.Shapes.AddFreeform(0, 0, 0, 0, 300, 200, new ShapePath[] { path });

                // Retrieve the shape we just added (first shape in the collection)
                Shape shape = sheet.Shapes[0];

                // Cast the geometry to CustomGeometry to access its paths
                if (shape.Geometry is not CustomGeometry geometry)
                {
                    Console.WriteLine("The shape does not contain custom geometry.");
                    return;
                }

                // Build a serializable object that represents the shape's path data
                var shapeData = new
                {
                    Paths = ExtractPathsInfo(geometry.Paths)
                };

                // Serialize the object to formatted JSON
                string json = JsonSerializer.Serialize(shapeData, new JsonSerializerOptions { WriteIndented = true });
                Console.WriteLine(json);

                // Save the workbook
                string outputPath = "ShapePathJsonDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Helper method to convert ShapePathCollection into a list of plain objects
        private static List<object> ExtractPathsInfo(ShapePathCollection paths)
        {
            var result = new List<object>();
            foreach (ShapePath p in paths)
            {
                var segmentList = new List<object>();
                foreach (ShapeSegmentPath segment in p.PathSegementList)
                {
                    var pointList = new List<object>();
                    foreach (ShapePathPoint pt in segment.Points)
                    {
                        pointList.Add(new { X = pt.XPixel, Y = pt.YPixel });
                    }

                    segmentList.Add(new
                    {
                        Type = segment.Type.ToString(),
                        Points = pointList
                    });
                }

                result.Add(new { Segments = segmentList });
            }
            return result;
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            SerializeShapePathDemo.Run();
        }
    }
}