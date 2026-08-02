// Title: Aspose.Cells .NET – Extract all points from a non‑primitive shape’s path segments
// Description: This example creates a workbook, adds a custom (NotPrimitive) AutoShape, casts its Geometry to CustomGeometry, and iterates through every ShapePath, ShapeSegmentPath, and ShapePathPoint. It prints the XPixel and YPixel coordinates of each point and saves the workbook, demonstrating how to read detailed geometry data from non‑primitive shapes in Aspose.Cells for .NET.
// Keywords: Aspose.Cells | CustomGeometry | ShapePath | ShapeSegmentPath | ShapePathPoint | non‑primitive shape | extract shape coordinates | C# | .NET | XPixel | YPixel | path segment points | shape geometry extraction
// Common Searches: how to read custom geometry points from a non‑primitive shape in Aspose.Cells | iterate ShapePath segments to get coordinates using Aspose.Cells C# | retrieve XPixel and YPixel values of shape path points | Aspose.Cells extract points from custom AutoShape | C# sample for accessing shape geometry in Excel workbook
// Developer Intent: Obtain the XPixel/YPixel coordinates of every point in each segment of a shape that uses custom geometry.
// Use Cases: Validate a custom shape by comparing extracted point coordinates with design specifications. | Convert shape path data to SVG or other vector formats for web rendering. | Perform geometric analysis such as collision detection or measurement within diagramming tools.
// AI Prompts: Write C# code that saves all extracted ShapePathPoint XPixel and YPixel values to a CSV file using Aspose.Cells. | Create a method that returns a List<PointF> containing every point from a non‑primitive shape’s CustomGeometry. | Explain how to modify the sample to include shape rotation and obtain transformed point coordinates.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapePathExtraction
{
    // This example creates a workbook, adds a custom (NotPrimitive) AutoShape, casts its Geometry to CustomGeometry, and iterates through every ShapePath, ShapeSegmentPath, and ShapePathPoint. It prints the XPixel and YPixel coordinates of each point and saves the workbook, demonstrating how to read detailed geometry data from non‑primitive shapes in Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a non‑primitive (custom) shape to the worksheet
            // Parameters: shape type, upper‑left row, upper‑left column, top offset, left offset, height, width
            Shape shape = worksheet.Shapes.AddAutoShape(
                AutoShapeType.NotPrimitive, // non‑primitive shape
                2, 2,                       // row, column
                0, 0,                       // top, left offsets
                200, 200);                  // height, width

            // Cast the shape's geometry to CustomGeometry to access its paths
            CustomGeometry geometry = shape.Geometry as CustomGeometry;
            if (geometry == null)
            {
                Console.WriteLine("The shape does not have custom geometry.");
                workbook.Save("NoCustomGeometry.xlsx");
                return;
            }

            // Iterate through each ShapePath in the geometry
            for (int p = 0; p < geometry.Paths.Count; p++)
            {
                ShapePath path = geometry.Paths[p];
                Console.WriteLine($"Path {p} - Segment count: {path.PathSegementList.Count}");

                // Iterate through each segment in the current path
                for (int s = 0; s < path.PathSegementList.Count; s++)
                {
                    ShapeSegmentPath segment = path.PathSegementList[s];
                    Console.WriteLine($"  Segment {s} - Type: {segment.Type}, Points: {segment.Points.Count}");

                    // Iterate through each point in the segment and display its coordinates
                    for (int pt = 0; pt < segment.Points.Count; pt++)
                    {
                        ShapePathPoint point = segment.Points[pt];
                        // Use pixel properties for readability
                        Console.WriteLine($"    Point {pt}: X={point.XPixel}, Y={point.YPixel}");
                    }
                }
            }

            // Save the workbook (the shape itself is saved for verification)
            workbook.Save("ShapePathSegmentsPointsDemo.xlsx");
        }
    }
}
