// Title: Validate shape positions and dimensions before and after SpreadsheetML export/import using Aspose.Cells for .NET
// Common Searches: how to compare shape coordinates before and after saving a workbook as XML with Aspose.Cells | C# code to detect shape size changes after SpreadsheetML round‑trip in Aspose.Cells | verify that shapes retain top left height width after exporting to SpreadsheetML using Aspose.Cells | Aspose.Cells shape geometry tolerance check after XML import | list missing or new shapes after loading a workbook from XML in C#
// Tags: shape geometry comparison Aspose.Cells | SpreadsheetML roundtrip shape position validation | extract shape dimensions worksheet Aspose.Cells | tolerance based shape coordinate check C# | detect missing shapes after XML import Aspose.Cells

using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example extracts each shape's Top, Left, Height, and Width from the first worksheet, saves the workbook as SpreadsheetML (XML), reloads it, and then compares the original and imported geometries with a small tolerance, reporting any missing, new, or altered shapes.
class GeometryComparer
{
    // Structure to hold shape geometry
    private class ShapeGeometry
    {
        public string Name { get; set; }
        public double Top { get; set; }
        public double Left { get; set; }
        public double Height { get; set; }
        public double Width { get; set; }

        public override bool Equals(object obj)
        {
            if (obj is ShapeGeometry other)
            {
                // Allow a small tolerance for floating point differences
                const double tolerance = 0.001;
                return Name == other.Name &&
                       Math.Abs(Top - other.Top) < tolerance &&
                       Math.Abs(Left - other.Left) < tolerance &&
                       Math.Abs(Height - other.Height) < tolerance &&
                       Math.Abs(Width - other.Width) < tolerance;
            }
            return false;
        }

        public override int GetHashCode()
        {
            return HashCode.Combine(Name, Top, Left, Height, Width);
        }
    }

    // Extract geometry of all shapes in the first worksheet
    private static List<ShapeGeometry> ExtractGeometries(Workbook wb)
    {
        var geometries = new List<ShapeGeometry>();
        Worksheet sheet = wb.Worksheets[0];
        foreach (Shape shape in sheet.Shapes)
        {
            geometries.Add(new ShapeGeometry
            {
                Name = shape.Name,
                Top = shape.Top,
                Left = shape.Left,
                Height = shape.Height,
                Width = shape.Width
            });
        }
        return geometries;
    }

    // Compare two geometry collections and report mismatches
    private static void CompareGeometries(List<ShapeGeometry> before, List<ShapeGeometry> after)
    {
        var beforeDict = new Dictionary<string, ShapeGeometry>();
        foreach (var g in before) beforeDict[g.Name] = g;

        var afterDict = new Dictionary<string, ShapeGeometry>();
        foreach (var g in after) afterDict[g.Name] = g;

        Console.WriteLine("=== Geometry Comparison Results ===");

        // Check for missing or changed shapes
        foreach (var kvp in beforeDict)
        {
            if (!afterDict.ContainsKey(kvp.Key))
            {
                Console.WriteLine($"Shape '{kvp.Key}' missing after import.");
                continue;
            }

            if (!kvp.Value.Equals(afterDict[kvp.Key]))
            {
                Console.WriteLine($"Shape '{kvp.Key}' geometry changed:");
                Console.WriteLine($"  Before - Top:{kvp.Value.Top}, Left:{kvp.Value.Left}, Height:{kvp.Value.Height}, Width:{kvp.Value.Width}");
                var a = afterDict[kvp.Key];
                Console.WriteLine($"  After  - Top:{a.Top}, Left:{a.Left}, Height:{a.Height}, Width:{a.Width}");
            }
        }

        // Check for new shapes added after import
        foreach (var kvp in afterDict)
        {
            if (!beforeDict.ContainsKey(kvp.Key))
            {
                Console.WriteLine($"New shape '{kvp.Key}' appeared after import.");
            }
        }

        Console.WriteLine("=== Comparison Completed ===");
    }

    static void Main()
    {
        // Load the original workbook (replace with your file path)
        Workbook originalWb = new Workbook("Original.xlsx");

        // Extract geometry before XML export
        List<ShapeGeometry> beforeGeometries = ExtractGeometries(originalWb);

        // Save the workbook as XML (SpreadsheetML)
        originalWb.Save("Exported.xml", SaveFormat.Xml);

        // Load the workbook back from the XML file
        Workbook importedWb = new Workbook("Exported.xml");

        // Extract geometry after XML import
        List<ShapeGeometry> afterGeometries = ExtractGeometries(importedWb);

        // Compare and output results
        CompareGeometries(beforeGeometries, afterGeometries);
    }
}
