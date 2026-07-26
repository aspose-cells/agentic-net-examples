// Title: Validate Shape Geometry After ImportXml with Aspose.Cells for .NET
// Description: A C# example that creates a workbook, adds a rectangle shape, records its Geometry, ShapeAdjustValues count, and Paths count, imports an XML file, re‑examines the same shape, and reports whether the geometry remains unchanged before saving the file.
// Keywords: Aspose.Cells ImportXml | shape geometry consistency | C# rectangle shape | ShapeAdjustValues | Paths count | geometry comparison | .NET spreadsheet API
// Common Searches: Aspose.Cells verify shape geometry after ImportXml | compare shape adjust values before and after XML import | does ImportXml affect rectangle shape paths | check shape geometry preservation in Aspose.Cells | C# test shape consistency after importing XML
// Developer Intent: Confirm that importing XML data does not modify the geometry of existing shapes in a workbook.
// Use Cases: Automated regression test to ensure shape dimensions stay intact after data import | Audit workflow that logs any geometry changes when XML is added to a spreadsheet | Unit test for validating that ImportXml preserves ShapeAdjustValues and path definitions
// AI Prompts: Generate C# code using Aspose.Cells that compares a shape's Geometry, ShapeAdjustValues count, and Paths count before and after ImportXml and outputs the result. | Write a .NET unit test that asserts a rectangle shape's adjust values and path count are unchanged after Workbook.ImportXml is called. | Explain how the Geometry, ShapeAdjustValues, and Paths properties behave during an ImportXml operation in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsGeometryComparison
{
    // A C# example that creates a workbook, adds a rectangle shape, records its Geometry, ShapeAdjustValues count, and Paths count, imports an XML file, re‑examines the same shape, and reports whether the geometry remains unchanged before saving the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            Shape rect = sheet.Shapes.AddRectangle(1, 1, 0, 0, 200, 100);

            // Access the geometry of the shape before XML import
            Geometry geometryBefore = rect.Geometry;
            int adjustCountBefore = geometryBefore.ShapeAdjustValues.Count;
            int pathCountBefore = rect.Paths != null ? rect.Paths.Count : 0;

            // Output initial geometry details
            Console.WriteLine("Before ImportXml:");
            Console.WriteLine($"Adjust values count: {adjustCountBefore}");
            Console.WriteLine($"Paths count: {pathCountBefore}");

            // Create a simple XML file (does not affect shapes)
            string xmlContent = @"<Root><Item>Value</Item></Root>";
            string xmlPath = "data.xml";
            File.WriteAllText(xmlPath, xmlContent);

            // Import the XML data into the workbook (starting at cell A1)
            workbook.ImportXml(xmlPath, "Sheet1", 0, 0);

            // Access the geometry of the same shape after XML import
            Geometry geometryAfter = rect.Geometry;
            int adjustCountAfter = geometryAfter.ShapeAdjustValues.Count;
            int pathCountAfter = rect.Paths != null ? rect.Paths.Count : 0;

            // Output geometry details after import
            Console.WriteLine("\nAfter ImportXml:");
            Console.WriteLine($"Adjust values count: {adjustCountAfter}");
            Console.WriteLine($"Paths count: {pathCountAfter}");

            // Compare and report consistency
            bool isAdjustCountEqual = adjustCountBefore == adjustCountAfter;
            bool isPathCountEqual = pathCountBefore == pathCountAfter;

            Console.WriteLine("\nComparison Result:");
            Console.WriteLine($"Adjust values count unchanged: {isAdjustCountEqual}");
            Console.WriteLine($"Paths count unchanged: {isPathCountEqual}");

            // Save the workbook (demonstrating standard save operation)
            workbook.Save("GeometryComparisonResult.xlsx");
        }
    }
}
