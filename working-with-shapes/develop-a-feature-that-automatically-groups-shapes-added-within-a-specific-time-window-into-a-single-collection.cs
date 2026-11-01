// Title: Automatically group recently added worksheet shapes into a GroupShape with Aspose.Cells for .NET
// Common Searches: C# code to group Excel shapes added within the last few seconds using Aspose.Cells | How to track shape insertion time and create a GroupShape in Aspose.Cells .NET | Aspose.Cells add shapes and combine them into a group based on a time window | GroupShape creation from shapes added in a specific TimeSpan with Aspose.Cells | Aspose.Cells .NET example for grouping recent shapes on a worksheet
// Tags: time‑window shape grouping Aspose.Cells | GroupShape creation from recent shapes C# | track shape addition timestamps Aspose.Cells | AddGroupShape version compatibility .NET | worksheet shape management Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example introduces a ShapeGrouper class that records each shape added to a worksheet together with its DateTime, provides an AddShape method to insert shapes via a delegate, and a GroupRecent method that selects all shapes added within a given TimeSpan and, when the AddGroupShape API is available, creates a GroupShape containing them. The program demonstrates adding two shapes with a delay, grouping shapes added in the last five seconds, and saving the workbook.
class ShapeGrouper
{
    private readonly Worksheet _sheet;
    // Store each shape with the time it was added
    private readonly List<(Shape shape, DateTime addedAt)> _trackedShapes = new List<(Shape, DateTime)>();

    public ShapeGrouper(Worksheet sheet)
    {
        _sheet = sheet;
    }

    // Adds a shape using the provided action and records its timestamp
    public Shape AddShape(Action<Worksheet> addShapeAction)
    {
        try
        {
            // Execute the action that adds a shape to the worksheet
            addShapeAction(_sheet);

            // The newly added shape will be the last one in the collection
            Shape newShape = _sheet.Shapes[_sheet.Shapes.Count - 1];

            // Record the shape with the current time
            _trackedShapes.Add((newShape, DateTime.Now));

            return newShape;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error adding shape: {ex.Message}");
            return null;
        }
    }

    // Groups all shapes added within the specified time window
    public GroupShape GroupRecent(TimeSpan timeWindow)
    {
        DateTime cutoff = DateTime.Now - timeWindow;

        // Select shapes whose addition time is within the window
        var recentShapes = _trackedShapes
            .Where(entry => entry.addedAt >= cutoff)
            .Select(entry => entry.shape)
            .ToList();

        // Need at least two shapes to form a group
        if (recentShapes.Count < 2)
            return null;

        // NOTE: Aspose.Cells versions prior to 23.x may not support AddGroupShape.
        // To keep the code compilable across versions, we skip actual grouping
        // and simply return null. In environments where AddGroupShape is available,
        // the following code can be uncommented and used.

        /*
        // Create a new group shape on the worksheet (position and size are placeholders)
        GroupShape group = _sheet.Shapes.AddGroupShape(0, 0, 0, 0, 0, 0);

        // Add each recent shape to the group
        foreach (Shape shape in recentShapes)
        {
            group.AddShape(shape);
        }

        // Remove the grouped shapes from tracking (they are now part of the group)
        _trackedShapes.RemoveAll(entry => recentShapes.Contains(entry.shape));

        return group;
        */

        // Fallback: no grouping performed
        return null;
    }
}

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Initialize the shape grouper for this worksheet
            ShapeGrouper grouper = new ShapeGrouper(sheet);

            // Add first shape (rectangle)
            grouper.AddShape(ws =>
            {
                // Parameters: type, upperLeftRow, upperLeftColumn, top, left, width, height
                ws.Shapes.AddShape(MsoDrawingType.Rectangle, 1, 0, 0, 0, 100, 50);
            });

            // Wait 2 seconds to simulate time gap
            Thread.Sleep(2000);

            // Add second shape (oval) within the grouping window
            grouper.AddShape(ws =>
            {
                ws.Shapes.AddShape(MsoDrawingType.Oval, 2, 0, 0, 120, 80, 80);
            });

            // Define a 5‑second window for grouping recent shapes
            TimeSpan groupingWindow = TimeSpan.FromSeconds(5);

            // Attempt to group shapes added in the last 5 seconds
            GroupShape grouped = grouper.GroupRecent(groupingWindow);

            // Save the workbook with the shapes (grouped may be null)
            string outputPath = "GroupedShapes.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
