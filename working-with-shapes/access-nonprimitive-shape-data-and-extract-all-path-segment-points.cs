// Title: Extract custom geometry path segment points from Excel worksheet shapes using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells read custom geometry points from Excel shapes | How to get path segment coordinates of a shape in an .xlsx file with Aspose.Cells | Retrieve non‑primitive shape data (CustomGeometry) using Aspose.Cells .NET | Iterate shape.PathList segments to extract point coordinates in Aspose.Cells | Extract shape path points from a workbook with Aspose.Cells for .NET
// Tags: extract shape custom geometry points Aspose.Cells | iterate shape pathlist segments C# | read shape segment coordinates Aspose.Cells .NET | custom geometry extraction from Excel shapes | shape point coordinates Aspose.Cells

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Reflection;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an Excel workbook, loops through all shapes on the first worksheet, and uses dynamic reflection to access each shape's CustomGeometry. It then traverses the PathList, Segments, and Points collections, converts each point to a PointF, aggregates them, and prints the X/Y coordinates while handling missing files and runtime errors.
class ExtractShapePathPoints
{
    static void Main()
    {
        const string inputFile = "input.xlsx";

        // Verify that the input workbook exists to avoid FileNotFoundException
        if (!File.Exists(inputFile))
        {
            Console.WriteLine($"Error: The file \"{inputFile}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputFile);
            Worksheet sheet = workbook.Worksheets[0];

            // Store extracted points
            List<PointF> allPoints = new List<PointF>();

            // Iterate through all shapes on the worksheet
            foreach (Shape shape in sheet.Shapes)
            {
                // Use dynamic access to handle possible version differences
                dynamic dynShape = shape;

                // Check if the shape exposes a CustomGeometry property at runtime
                if (HasProperty(dynShape, "CustomGeometry"))
                {
                    try
                    {
                        dynamic customGeometry = dynShape.CustomGeometry;
                        if (customGeometry != null && HasProperty(customGeometry, "PathList"))
                        {
                            foreach (dynamic path in customGeometry.PathList)
                            {
                                if (HasProperty(path, "Segments"))
                                {
                                    foreach (dynamic segment in path.Segments)
                                    {
                                        if (HasProperty(segment, "Points"))
                                        {
                                            foreach (dynamic pt in segment.Points)
                                            {
                                                // Convert to PointF (Aspose may return PointF already)
                                                float x = Convert.ToSingle(pt.X);
                                                float y = Convert.ToSingle(pt.Y);
                                                allPoints.Add(new PointF(x, y));
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        // Handle any unexpected issues while processing a shape's geometry
                        Console.WriteLine($"Warning: Could not process shape ID {shape.Id}. {ex.Message}");
                    }
                }
            }

            // Output the extracted points
            Console.WriteLine("Extracted Path Segment Points:");
            foreach (PointF pt in allPoints)
            {
                Console.WriteLine($"X = {pt.X}, Y = {pt.Y}");
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Helper to determine if a dynamic object has a given property
    private static bool HasProperty(object obj, string propertyName)
    {
        if (obj == null) return false;
        return obj.GetType().GetProperty(propertyName, BindingFlags.Public | BindingFlags.Instance) != null;
    }
}
