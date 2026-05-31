using System;
using System.Collections.Generic;
using System.Threading;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapeGroupingDemo
{
    class Program
    {
        // Holds a shape together with the time it was added
        private static readonly List<(Shape shape, DateTime addedTime)> shapeLog = new List<(Shape, DateTime)>();

        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            ShapeCollection shapes = worksheet.Shapes;

            // Add shapes at different moments (simulated with Thread.Sleep)
            AddRectangle(shapes, 2, 0, 2, 0, 80, 60);   // time = T0
            Thread.Sleep(1500); // 1.5 seconds later
            AddOval(shapes, 6, 0, 2, 0, 80, 60);        // time = T0 + 1.5s
            Thread.Sleep(3000); // 3 seconds later
            AddRectangle(shapes, 10, 0, 2, 0, 80, 60); // time = T0 + 4.5s
            Thread.Sleep(1000); // 1 second later
            AddOval(shapes, 14, 0, 2, 0, 80, 60);       // time = T0 + 5.5s

            // Define a time window (e.g., 3 seconds) and group shapes added within that window
            TimeSpan window = TimeSpan.FromSeconds(3);
            GroupRecentShapes(shapes, window);

            // Save the workbook
            workbook.Save("GroupedShapesDemo.xlsx");
        }

        // Adds a rectangle shape and records its addition time
        private static void AddRectangle(ShapeCollection shapes, int upperLeftRow, int top, int upperLeftColumn, int left, int height, int width)
        {
            Shape rect = shapes.AddRectangle(upperLeftRow, top, upperLeftColumn, left, height, width);
            rect.Name = $"Rect_{shapes.Count}";
            shapeLog.Add((rect, DateTime.UtcNow));
        }

        // Adds an oval shape and records its addition time
        private static void AddOval(ShapeCollection shapes, int upperLeftRow, int top, int upperLeftColumn, int left, int height, int width)
        {
            Shape oval = shapes.AddOval(upperLeftRow, top, upperLeftColumn, left, height, width);
            oval.Name = $"Oval_{shapes.Count}";
            shapeLog.Add((oval, DateTime.UtcNow));
        }

        // Groups all shapes that were added within the specified time window relative to the most recent addition
        private static void GroupRecentShapes(ShapeCollection shapes, TimeSpan window)
        {
            if (shapeLog.Count == 0) return;

            // Determine the cutoff time: shapes added after (most recent time - window) are eligible
            DateTime mostRecent = shapeLog[shapeLog.Count - 1].addedTime;
            DateTime cutoff = mostRecent - window;

            // Collect eligible shapes
            List<Shape> toGroup = new List<Shape>();
            foreach (var entry in shapeLog)
            {
                if (entry.addedTime >= cutoff && !entry.shape.IsInGroup)
                {
                    toGroup.Add(entry.shape);
                }
            }

            // If we have at least two shapes, group them
            if (toGroup.Count >= 2)
            {
                GroupShape group = shapes.Group(toGroup.ToArray());
                group.Name = $"Group_{DateTime.UtcNow:yyyyMMdd_HHmmss}";
            }
        }
    }
}