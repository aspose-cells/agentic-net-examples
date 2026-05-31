using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsZOrderDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add shapes with a custom metadata field encoded in the shape's Name.
            // Format: "ShapeName|Importance=Level"
            Shape shapeA = worksheet.Shapes.AddRectangle(10, 10, 100, 100, 0, 0);
            shapeA.Name = "ShapeA|Importance=1";

            Shape shapeB = worksheet.Shapes.AddRectangle(50, 50, 100, 100, 0, 0);
            shapeB.Name = "ShapeB|Importance=3";

            Shape shapeC = worksheet.Shapes.AddRectangle(90, 90, 100, 100, 0, 0);
            shapeC.Name = "ShapeC|Importance=2";

            // Collect all shapes and their importance levels
            List<(Shape shape, int importance)> shapeInfo = new List<(Shape, int)>();
            foreach (Shape shp in worksheet.Shapes)
            {
                int importance = 0; // default importance
                // Parse the custom metadata from the Name property
                // Expected pattern: "...|Importance=Level"
                string[] parts = shp.Name.Split('|');
                foreach (string part in parts)
                {
                    if (part.StartsWith("Importance=", StringComparison.OrdinalIgnoreCase))
                    {
                        string value = part.Substring("Importance=".Length);
                        int.TryParse(value, out importance);
                        break;
                    }
                }
                shapeInfo.Add((shp, importance));
            }

            // Sort shapes by importance descending (higher importance -> front)
            shapeInfo.Sort((x, y) => y.importance.CompareTo(x.importance));

            // Assign ZOrderPosition so that the most important shape gets the highest position
            // ZOrderPosition 0 is the backmost; higher values are closer to the front.
            for (int i = 0; i < shapeInfo.Count; i++)
            {
                // Backmost shape gets position 0, next gets 1, etc.
                shapeInfo[i].shape.ZOrderPosition = i;
            }

            // Save the workbook
            workbook.Save("ShapesZOrderByImportance.xlsx");
        }
    }
}