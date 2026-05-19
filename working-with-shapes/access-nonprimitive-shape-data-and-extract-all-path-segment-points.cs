using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class ExtractShapePathPoints
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Build a simple freeform shape (rectangle) using ShapePath commands
        ShapePath shapePath = new ShapePath();
        shapePath.MoveTo(10, 10);
        shapePath.LineTo(200, 10);
        shapePath.LineTo(200, 100);
        shapePath.LineTo(10, 100);
        shapePath.Close();

        // Add the freeform shape to the worksheet
        worksheet.Shapes.AddFreeform(1, 0, 1, 0, 300, 200, new ShapePath[] { shapePath });

        // Retrieve the shape we just added (first shape in the collection)
        Shape shape = worksheet.Shapes[0];

        // Cast the shape's geometry to CustomGeometry to access path information
        CustomGeometry geometry = shape.Geometry as CustomGeometry;
        if (geometry != null)
        {
            // Iterate through each ShapePath in the geometry
            for (int pathIndex = 0; pathIndex < geometry.Paths.Count; pathIndex++)
            {
                ShapePath path = geometry.Paths[pathIndex];

                // Iterate through each segment in the current ShapePath
                for (int segIndex = 0; segIndex < path.PathSegementList.Count; segIndex++)
                {
                    ShapeSegmentPath segment = path.PathSegementList[segIndex];
                    Console.WriteLine($"Path {pathIndex}, Segment {segIndex}, Type: {segment.Type}, Points: {segment.Points.Count}");

                    // Iterate through all points in the current segment
                    for (int ptIndex = 0; ptIndex < segment.Points.Count; ptIndex++)
                    {
                        ShapePathPoint point = segment.Points[ptIndex];
                        // Use pixel coordinates for readability
                        Console.WriteLine($"  Point {ptIndex}: X = {point.XPixel}, Y = {point.YPixel}");
                    }
                }
            }
        }

        // Save the workbook
        workbook.Save("ExtractShapePathPoints.xlsx");
    }
}