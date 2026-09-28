// Title: How to automatically group Excel shapes added within a specific time window using Aspose.Cells for .NET
// AI Prompts: Write a C# class that records the timestamp of each shape added to an Aspose.Cells worksheet and groups shapes whose timestamps fall inside a given TimeSpan. | Extend the ShapeGrouper to expose a method that returns collections of shapes grouped by a configurable time interval without persisting the workbook. | Add logging and error handling to skip grouping when the Aspose.Cells library does not provide a GroupShape API.
// Common Searches: aspnet shape grouping based on insertion time Aspose.Cells | c# group Excel shapes added within 5 seconds using Aspose | track shape addition timestamps in Aspose.Cells workbook | how to create collections of shapes by time interval in Aspose.Cells .NET
// Tags: group shapes by time interval Aspose.Cells | record shape insertion time C# | time‑based shape collection Aspose.Cells | Aspose.Cells shape grouping limitation | shape grouping using TimeSpan .NET

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example defines a ShapeGrouper class that logs the DateTime each shape is added to a worksheet, orders the shapes chronologically, and creates logical groups for shapes whose addition times are within a specified TimeSpan. Because the current Aspose.Cells API lacks a visual GroupShape method, the CreateGroup placeholder does not modify the workbook, but the logic can be extended when the API becomes available.
class ShapeGrouper
{
    private Workbook _workbook;
    private Worksheet _sheet;
    private readonly Dictionary<Shape, DateTime> _shapeTimes = new Dictionary<Shape, DateTime>();

    public ShapeGrouper()
    {
        // Initialize a new workbook and get the first worksheet
        _workbook = new Workbook();
        _sheet = _workbook.Worksheets[0];
    }

    public void Load(string path)
    {
        try
        {
            if (!File.Exists(path))
                throw new FileNotFoundException($"Template file not found: {path}");

            _workbook = new Workbook(path);
            _sheet = _workbook.Worksheets[0];
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading workbook: {ex.Message}");
            throw;
        }
    }

    // Records the addition time of a shape
    public void AddShape(Shape shape)
    {
        _shapeTimes[shape] = DateTime.Now;
    }

    // Sample method that adds shapes with delays to simulate a time window
    public void AddSampleShapes()
    {
        try
        {
            var shape1 = _sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 1, 1, 0, 0, 100, 50);
            AddShape(shape1);
            Thread.Sleep(2000); // 2 seconds

            var shape2 = _sheet.Shapes.AddShape(MsoDrawingType.Oval, 2, 1, 0, 0, 80, 80);
            AddShape(shape2);
            Thread.Sleep(3000); // 3 seconds

            var shape3 = _sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 3, 1, 0, 0, 120, 60);
            AddShape(shape3);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error adding sample shapes: {ex.Message}");
            throw;
        }
    }

    // Groups shapes whose addition times fall within the specified time window
    public void GroupShapesWithin(TimeSpan window)
    {
        var ordered = _shapeTimes.OrderBy(kv => kv.Value).ToList();

        List<Shape> currentGroup = new List<Shape>();
        DateTime? groupStart = null;

        foreach (var kv in ordered)
        {
            if (groupStart == null)
            {
                groupStart = kv.Value;
                currentGroup.Add(kv.Key);
                continue;
            }

            if (kv.Value - groupStart <= window)
            {
                currentGroup.Add(kv.Key);
            }
            else
            {
                if (currentGroup.Count > 1)
                    CreateGroup(currentGroup);

                currentGroup = new List<Shape> { kv.Key };
                groupStart = kv.Value;
            }
        }

        if (currentGroup.Count > 1)
            CreateGroup(currentGroup);
    }

    // Placeholder for grouping logic.
    // Aspose.Cells version used does not expose GroupShape APIs, so this method currently
    // leaves the shapes unchanged. Adjust as needed when a compatible API is available.
    private void CreateGroup(List<Shape> shapes)
    {
        // No grouping operation performed due to missing API support.
        // This method can be expanded to implement visual grouping when supported.
    }

    public void Save(string path)
    {
        try
        {
            _workbook.Save(path);
            Console.WriteLine($"Workbook saved to {path}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error saving workbook: {ex.Message}");
            throw;
        }
    }
}

class Program
{
    static void Main()
    {
        try
        {
            var grouper = new ShapeGrouper();

            // Add shapes with simulated time intervals
            grouper.AddSampleShapes();

            // Group shapes added within a 5‑second window
            grouper.GroupShapesWithin(TimeSpan.FromSeconds(5));

            // Save the resulting workbook
            grouper.Save("GroupedShapes.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unhandled exception: {ex.Message}");
        }
    }
}
