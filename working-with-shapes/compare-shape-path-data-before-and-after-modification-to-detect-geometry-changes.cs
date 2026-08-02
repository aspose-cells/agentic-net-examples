// Title: Compare ShapePath data to detect geometry changes in Aspose.Cells for .NET
// Description: This example creates a workbook, adds a NotPrimitive autoshape, records its ShapePathCollection as readable strings, modifies the shape by adding a line segment, captures the updated path data, and programmatically determines whether the geometry has changed before saving the file.
// Keywords: Aspose.Cells ShapePath comparison | detect shape geometry changes .NET | custom autoshape path data | ShapePathCollection snapshot | C# Aspose.Cells shape editing
// Common Searches: compare shape path data Aspose.Cells | detect changes in custom autoshape geometry | Aspose.Cells get shape segment points | track shape modifications in spreadsheet | C# check if shape geometry changed
// Developer Intent: Identify whether a shape’s geometry has been altered by comparing its path data before and after programmatic edits.
// Use Cases: Validate that automated shape edits preserve intended geometry. | Implement change‑tracking for custom shapes in spreadsheet generation pipelines. | Create unit tests that verify specific path segment modifications.
// AI Prompts: Write C# code using Aspose.Cells to capture a ShapePathCollection, modify the shape, and compare the before/after data to detect geometry changes. | Show how to log detailed differences between original and modified ShapePath segment strings. | Explain how to extend the comparison to handle multiple ShapePath objects and report added, removed, or altered segments.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapeGeometryComparison
{
    // This example creates a workbook, adds a NotPrimitive autoshape, records its ShapePathCollection as readable strings, modifies the shape by adding a line segment, captures the updated path data, and programmatically determines whether the geometry has changed before saving the file.
    public class CompareShapePathData
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a NotPrimitive autoshape (custom geometry) to the worksheet
                // Parameters: shape type, upper left row, column, upper left offset X/Y, width, height
                Shape shape = worksheet.Shapes.AddAutoShape(AutoShapeType.NotPrimitive, 0, 0, 0, 0, 200, 200);

                // Access the shape's path collection
                ShapePathCollection paths = shape.Paths;

                // Helper method to capture the current path data as a list of strings
                List<string> CapturePathData(ShapePathCollection collection)
                {
                    var data = new List<string>();
                    for (int i = 0; i < collection.Count; i++)
                    {
                        ShapePath path = collection[i];
                        // Record each segment type and its points
                        foreach (ShapeSegmentPath segment in path.PathSegementList)
                        {
                            // Build a string representation of the segment
                            string segmentInfo = $"Path{i}_Segment{segment.Type}";
                            int pointIndex = 0;
                            foreach (ShapePathPoint point in segment.Points)
                            {
                                // Use non‑obsolete pixel properties
                                segmentInfo += $"_P{pointIndex}({point.XPixel},{point.YPixel})";
                                pointIndex++;
                            }
                            data.Add(segmentInfo);
                        }
                    }
                    return data;
                }

                // Capture the initial geometry data
                List<string> beforeData = CapturePathData(paths);

                // Modify the geometry: add a new line to the first path
                if (paths.Count > 0)
                {
                    ShapePath firstPath = paths[0];
                    // Ensure the path has a starting point; if not, move to (0,0)
                    if (firstPath.PathSegementList.Count == 0)
                    {
                        firstPath.MoveTo(0, 0);
                    }
                    // Add a line segment
                    firstPath.LineTo(150, 150);
                }
                else
                {
                    // If there are no paths, create one and add a line
                    int newPathIndex = paths.Add();
                    ShapePath newPath = paths[newPathIndex];
                    newPath.MoveTo(0, 0);
                    newPath.LineTo(150, 150);
                }

                // Capture the geometry data after modification
                List<string> afterData = CapturePathData(paths);

                // Compare the two snapshots
                bool geometryChanged = false;
                if (beforeData.Count != afterData.Count)
                {
                    geometryChanged = true;
                }
                else
                {
                    for (int i = 0; i < beforeData.Count; i++)
                    {
                        if (!beforeData[i].Equals(afterData[i], StringComparison.Ordinal))
                        {
                            geometryChanged = true;
                            break;
                        }
                    }
                }

                // Output the result
                Console.WriteLine(geometryChanged
                    ? "Shape geometry has changed after modification."
                    : "Shape geometry remains unchanged.");

                // Save the workbook (required by lifecycle rule)
                string outputPath = "ShapeGeometryComparison.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            CompareShapePathData.Run();
        }
    }
}
