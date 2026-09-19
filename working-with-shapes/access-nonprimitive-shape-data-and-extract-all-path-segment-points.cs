// Title: How to extract coordinates of free‑form shape path segments from an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Write a C# program with Aspose.Cells that loads an .xlsx file, iterates all worksheets, finds free‑form shapes, and prints each path segment’s type together with the X and Y values of its points. | Generate C# code that uses reflection to obtain a shape’s FreeFormPath and its Segments collection in Aspose.Cells, then enumerates every point in each segment and outputs the coordinates.
// Common Searches: c# aspocells read freeform shape points from excel | aspocells extract path segment coordinates from shape | how to get FreeFormPath segments using Aspose.Cells .NET | enumerate shape points in Excel workbook with Aspose.Cells C# | reflection based access to non‑primitive shape data Aspose.Cells
// Tags: Aspose.Cells freeform shape path extraction C# | enumerate shape segment points .NET | reflection access FreeFormPath Aspose.Cells | read non‑primitive shape data Excel | extract path segment coordinates from .xlsx

using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an Excel workbook, walks through each worksheet and its shapes, uses reflection to retrieve the FreeFormPath of free‑form shapes, iterates the Segments collection, and prints the segment type along with the X/Y coordinates of every point in each segment.
class ShapePathExtractor
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Ensure the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        Workbook workbook = null;
        try
        {
            // Load the workbook
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load workbook: {ex.Message}");
            return;
        }

        // Iterate through all worksheets
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            ShapeCollection shapes = sheet.Shapes;

            // Iterate through each shape
            foreach (Shape shape in shapes)
            {
                // Use reflection to safely access FreeFormPath (may not be available in older versions)
                var freeFormPathProp = shape.GetType().GetProperty("FreeFormPath");
                if (freeFormPathProp == null) continue; // Not a free‑form shape

                var freeFormPath = freeFormPathProp.GetValue(shape);
                if (freeFormPath == null) continue; // No path data

                // Retrieve the Segments collection via reflection
                var segmentsProp = freeFormPath.GetType().GetProperty("Segments");
                var segments = segmentsProp?.GetValue(freeFormPath) as IEnumerable;
                if (segments == null) continue;

                Console.WriteLine($"Worksheet: {sheet.Name}, Shape Name: {shape.Name}");

                // Iterate through each segment
                foreach (var segment in segments)
                {
                    // Segment type
                    var typeProp = segment.GetType().GetProperty("Type");
                    var segType = typeProp?.GetValue(segment);
                    Console.WriteLine($"  Segment Type: {segType}");

                    // Points collection
                    var pointsProp = segment.GetType().GetProperty("Points");
                    var points = pointsProp?.GetValue(segment) as IEnumerable;
                    if (points == null) continue;

                    int pointIndex = 0;
                    foreach (var ptObj in points)
                    {
                        // Each point is a System.Drawing.PointF
                        var xProp = ptObj.GetType().GetProperty("X");
                        var yProp = ptObj.GetType().GetProperty("Y");
                        var x = xProp?.GetValue(ptObj);
                        var y = yProp?.GetValue(ptObj);
                        Console.WriteLine($"    Point {pointIndex}: X = {x}, Y = {y}");
                        pointIndex++;
                    }
                }
            }
        }

        // If modifications were made, you could save the workbook here.
        // workbook.Save("output.xlsx");
    }
}
