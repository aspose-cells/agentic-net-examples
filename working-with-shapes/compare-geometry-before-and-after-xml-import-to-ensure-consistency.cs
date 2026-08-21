// Title: Check Shape Geometry Consistency After Workbook.ImportXml with Aspose.Cells for .NET
// Description: Creates a workbook, adds a rectangle shape with an adjustment value, records the counts of ShapeAdjustValues and Paths, imports a simple XML file using Workbook.ImportXml, re‑examines the same shape, and reports whether the geometry (adjustments and paths) remains unchanged before saving the file.
// Keywords: Aspose.Cells | Workbook.ImportXml | shape geometry | ShapeAdjustValues | shape paths | .NET C# | drawing objects integrity | XML import validation
// Common Searches: Aspose.Cells verify shape geometry after ImportXml | C# check if shape adjustments change when importing XML | Workbook.ImportXml impact on drawing objects | compare shape paths before and after XML import Aspose
// Developer Intent: Confirm that a shape’s geometry—adjustment values and path collection—remains unchanged after calling Workbook.ImportXml on a workbook.
// Use Cases: Automated regression test to ensure ImportXml does not alter existing drawing objects. | Audit routine for batch XML imports that must preserve custom shape adjustments. | Debugging scenario where unexpected shape distortion is suspected after XML processing.
// AI Prompts: Generate C# code using Aspose.Cells that records a shape's Geometry.ShapeAdjustValues and Paths counts, runs Workbook.ImportXml, then compares and logs any differences. | Write an MSTest unit test that asserts a rectangle shape's adjustment values and path collection stay the same after Workbook.ImportXml is executed. | Explain how Workbook.ImportXml interacts with drawing objects in Aspose.Cells and provide best‑practice tips to keep shape geometry intact during XML imports.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a workbook, adds a rectangle shape with an adjustment value, records the counts of ShapeAdjustValues and Paths, imports a simple XML file using Workbook.ImportXml, re‑examines the same shape, and reports whether the geometry (adjustments and paths) remains unchanged before saving the file.
class CompareShapeGeometry
{
    static void Main()
    {
        // Create a new workbook and add a rectangle shape
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Shape shape = sheet.Shapes.AddRectangle(1, 1, 0, 0, 200, 100);

        // Add an adjustment value to the shape's geometry
        shape.Geometry.ShapeAdjustValues.Add("adj1", 0.5);

        // Capture geometry details before XML import
        int beforeAdjustCount = shape.Geometry.ShapeAdjustValues.Count;
        int beforePathCount = shape.Paths != null ? shape.Paths.Count : 0;
        Console.WriteLine($"Before Import - Adjust values: {beforeAdjustCount}, Paths: {beforePathCount}");

        // Create a simple XML file (does not affect shapes)
        string xmlPath = "sample.xml";
        File.WriteAllText(xmlPath, "<root></root>");

        // Import the XML into the workbook (using the provided ImportXml rule)
        workbook.ImportXml(xmlPath, sheet.Name, 0, 0);

        // Capture geometry details after XML import
        Shape shapeAfter = sheet.Shapes[0]; // same shape instance
        int afterAdjustCount = shapeAfter.Geometry.ShapeAdjustValues.Count;
        int afterPathCount = shapeAfter.Paths != null ? shapeAfter.Paths.Count : 0;
        Console.WriteLine($"After Import - Adjust values: {afterAdjustCount}, Paths: {afterPathCount}");

        // Compare and report consistency
        bool adjustSame = beforeAdjustCount == afterAdjustCount;
        bool pathsSame = beforePathCount == afterPathCount;
        Console.WriteLine($"Geometry consistency - Adjust values same? {adjustSame}, Paths same? {pathsSame}");

        // Save the workbook (optional)
        workbook.Save("GeometryComparison.xlsx");
    }
}
