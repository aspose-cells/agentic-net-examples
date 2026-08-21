// Title: Auto‑group shapes added within a time window using Aspose.Cells for .NET (C#)
// Description: This example creates a workbook, adds rectangle shapes with simulated delays, records each shape's timestamp, and automatically groups shapes whose addition times fall inside a configurable TimeSpan (e.g., 5 seconds) using the Shapes.Group method. Each GroupShape is named with the window start time and the workbook is saved as an Excel file.
// Keywords: Aspose.Cells | C# | .NET | ShapeCollection | GroupShape | time‑based grouping | Shapes.Group | Excel shape automation | timestamp grouping | auto group shapes
// Common Searches: Aspose.Cells group shapes by time interval | C# automatically group Excel shapes | How to create GroupShape based on addition timestamp in Aspose.Cells | Time‑window shape grouping .NET | Group multiple shapes after adding them in Excel using Aspose
// Developer Intent: Implement automatic grouping of worksheet shapes that are added within a defined time interval.
// Use Cases: Combine annotation shapes created in quick succession so they move as a single object. | Batch user‑drawn shapes in a reporting tool into a GroupShape for easier formatting and deletion. | Generate Excel diagrams where sequentially inserted shapes need to be logically grouped without manual effort. | Provide a drawing canvas in an application that auto‑clusters shapes drawn within a short period. | Simplify cleanup of temporary shapes by grouping them based on their creation time.
// AI Prompts: Write a reusable method that takes a ShapeCollection and a TimeSpan, groups shapes added within that interval into GroupShape objects, and returns the created groups. | Refactor the sample to support overlapping time windows and allow the grouping window to be passed as a parameter. | Create unit tests for the time‑based shape grouping logic using Aspose.Cells mock objects to verify correct grouping and naming. | Design a helper class that tracks shape addition timestamps and performs automatic grouping when the workbook is saved. | Generate documentation comments for the time‑window grouping feature following the Aspose.Cells coding standards.

using System;
using System.Collections.Generic;
using System.Threading;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapeGroupingDemo
{
    // This example creates a workbook, adds rectangle shapes with simulated delays, records each shape's timestamp, and automatically groups shapes whose addition times fall inside a configurable TimeSpan (e.g., 5 seconds) using the Shapes.Group method. Each GroupShape is named with the window start time and the workbook is saved as an Excel file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            ShapeCollection shapes = worksheet.Shapes;

            // List to keep track of each shape and the time it was added
            List<(Shape shape, DateTime addedTime)> addedShapes = new List<(Shape, DateTime)>();

            // Define a helper method to add a rectangle and record its timestamp
            void AddRectangle(int row, int col, int height, int width)
            {
                Shape shape = shapes.AddRectangle(row, 0, col, 0, height, width);
                addedShapes.Add((shape, DateTime.Now));
            }

            // Add shapes with delays to simulate different addition times
            AddRectangle(2, 2, 50, 100);          // Shape 1
            Thread.Sleep(2000);                  // 2 seconds later
            AddRectangle(6, 2, 50, 100);          // Shape 2 (within 5‑second window)
            Thread.Sleep(4000);                  // 4 seconds later (total 6 seconds from first)
            AddRectangle(10, 2, 50, 100);         // Shape 3 (outside the 5‑second window)

            // Define the time window for grouping (e.g., 5 seconds)
            TimeSpan groupingWindow = TimeSpan.FromSeconds(5);

            // Group shapes that were added within the same time window
            List<Shape> currentGroup = new List<Shape>();
            DateTime? windowStart = null;

            foreach (var (shape, addedTime) in addedShapes)
            {
                if (windowStart == null)
                {
                    // Start a new window
                    windowStart = addedTime;
                    currentGroup.Add(shape);
                }
                else if (addedTime - windowStart <= groupingWindow)
                {
                    // Still within the window, add to current group
                    currentGroup.Add(shape);
                }
                else
                {
                    // Window exceeded, create group if more than one shape
                    if (currentGroup.Count > 1)
                    {
                        Shape[] groupArray = currentGroup.ToArray();
                        GroupShape groupShape = shapes.Group(groupArray);
                        groupShape.Name = $"Group_{windowStart:HHmmss}";
                    }

                    // Start a new window with the current shape
                    currentGroup.Clear();
                    currentGroup.Add(shape);
                    windowStart = addedTime;
                }
            }

            // Handle the last accumulated group
            if (currentGroup.Count > 1)
            {
                Shape[] groupArray = currentGroup.ToArray();
                GroupShape groupShape = shapes.Group(groupArray);
                groupShape.Name = $"Group_{windowStart:HHmmss}";
            }

            // Save the workbook
            workbook.Save("GroupedShapesDemo.xlsx");
        }
    }
}
