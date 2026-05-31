using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    public class ShapeGeometryChangeDetection
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a NotPrimitive autoshape (custom geometry)
                Shape shape = worksheet.Shapes.AddAutoShape(AutoShapeType.NotPrimitive, 0, 0, 0, 0, 200, 200);

                // The Geometry property is read‑only but already contains a CustomGeometry instance for NotPrimitive shapes
                CustomGeometry customGeometry = shape.Geometry as CustomGeometry;
                if (customGeometry == null)
                {
                    Console.WriteLine("The shape does not support custom geometry.");
                    return;
                }

                ShapePathCollection paths = customGeometry.Paths;

                // Capture the geometry state before modification
                int beforePathCount = paths.Count;
                List<int> beforeSegmentCounts = new List<int>();
                for (int i = 0; i < beforePathCount; i++)
                {
                    beforeSegmentCounts.Add(paths[i].PathSegementList.Count);
                }

                // Modify the geometry: add a new path forming a rectangle
                int newPathIndex = paths.Add();
                ShapePath newPath = paths[newPathIndex];
                newPath.MoveTo(0, 0);
                newPath.LineTo(100, 0);
                newPath.LineTo(100, 100);
                newPath.LineTo(0, 100);
                newPath.Close();

                // Capture the geometry state after modification
                int afterPathCount = paths.Count;
                List<int> afterSegmentCounts = new List<int>();
                for (int i = 0; i < afterPathCount; i++)
                {
                    afterSegmentCounts.Add(paths[i].PathSegementList.Count);
                }

                // Detect changes by comparing path counts and segment counts
                bool geometryChanged = false;
                if (beforePathCount != afterPathCount)
                {
                    geometryChanged = true;
                }
                else
                {
                    for (int i = 0; i < beforePathCount; i++)
                    {
                        if (beforeSegmentCounts[i] != afterSegmentCounts[i])
                        {
                            geometryChanged = true;
                            break;
                        }
                    }
                }

                // Output the detection result
                Console.WriteLine("Geometry changed: " + geometryChanged);

                // Save the workbook
                string outputPath = "ShapeGeometryChangeDetection.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    internal class Program
    {
        private static void Main(string[] args)
        {
            ShapeGeometryChangeDetection.Run();
        }
    }
}