// Title: Extract Shape Path Segment Points from a Non‑Primitive AutoShape with Aspose.Cells for .NET
// Description: Creates a workbook, adds a custom (non‑primitive) AutoShape, accesses its CustomGeometry, builds a rectangle path, then iterates every ShapePath, ShapeSegmentPath, and ShapePathPoint to output XPixel/YPixel coordinates before saving the file.
// Keywords: Aspose.Cells | CustomGeometry | ShapePathSegmentList | non‑primitive AutoShape | C# shape coordinates | extract shape points | .NET vector geometry | XPixel YPixel
// Common Searches: read custom shape points Aspose.Cells C# | iterate ShapeSegmentPath list .NET | get vertex coordinates of AutoShape | extract path segment coordinates Aspose.Cells
// Developer Intent: Programmatically retrieve every coordinate that defines the segments of a non‑primitive AutoShape’s geometry.
// Use Cases: Export shape vertices to SVG, PDF, or other vector formats. | Validate shape geometry by comparing segment points against design specifications. | Reconstruct or modify shapes dynamically based on their existing point data.
// AI Prompts: Write C# code using Aspose.Cells to loop through all ShapeSegmentPath objects of a CustomGeometry shape and print each point’s XPixel and YPixel values. | Create a method that accepts a Shape instance and returns a collection of (X, Y) coordinates for every segment in a non‑primitive AutoShape. | Explain how to add new points to an existing ShapePathSegment after extracting the current points with Aspose.Cells.

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;

// Creates a workbook, adds a custom (non‑primitive) AutoShape, accesses its CustomGeometry, builds a rectangle path, then iterates every ShapePath, ShapeSegmentPath, and ShapePathPoint to output XPixel/YPixel coordinates before saving the file.
class ExtractShapeSegmentPoints
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a non‑primitive auto shape (e.g., a custom shape)
        Shape shape = worksheet.Shapes.AddAutoShape(
            AutoShapeType.NotPrimitive, // shape type
            1, 1,                       // top row, top offset
            0, 0,                       // left column, left offset
            300, 300);                  // height, width

        // Cast the shape's geometry to CustomGeometry to access paths
        CustomGeometry geometry = shape.Geometry as CustomGeometry;
        if (geometry == null)
        {
            Console.WriteLine("The shape does not contain custom geometry.");
            return;
        }

        // Ensure there is at least one path; create one if none exist
        if (geometry.Paths.Count == 0)
        {
            geometry.Paths.Add();
        }

        // Build a simple rectangle path for demonstration
        ShapePath path = geometry.Paths[0];
        path.MoveTo(0, 0);
        path.LineTo(10000, 0);
        path.LineTo(10000, 10000);
        path.LineTo(0, 10000);
        path.Close();

        // Iterate through all paths, their segment paths, and points
        for (int p = 0; p < geometry.Paths.Count; p++)
        {
            ShapePath curPath = geometry.Paths[p];
            Console.WriteLine($"Path {p} contains {curPath.PathSegementList.Count} segment(s).");

            for (int s = 0; s < curPath.PathSegementList.Count; s++)
            {
                ShapeSegmentPath segment = curPath.PathSegementList[s];
                Console.WriteLine($"  Segment {s} Type: {segment.Type}, Points: {segment.Points.Count}");

                for (int pt = 0; pt < segment.Points.Count; pt++)
                {
                    ShapePathPoint point = segment.Points[pt];
                    // Use pixel coordinates for readability
                    Console.WriteLine($"    Point {pt}: X = {point.XPixel}, Y = {point.YPixel}");
                }
            }
        }

        // Save the workbook
        workbook.Save("ExtractShapeSegmentPoints.xlsx");
    }
}
