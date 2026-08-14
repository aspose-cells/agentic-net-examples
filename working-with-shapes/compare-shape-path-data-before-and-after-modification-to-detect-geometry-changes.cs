// Title: Compare Shape Path Data Before and After Modification with Aspose.Cells for .NET
// Description: This example creates a workbook, adds a NotPrimitive autoshape with custom geometry, records a string signature of all ShapePath segments, inserts a new rectangular path, records a second signature, and compares the two signatures to determine whether the shape's geometry has changed. The workbook is then saved.
// Keywords: Aspose.Cells | C# | .NET | shape geometry | custom geometry | ShapePath | CustomGeometry | compare shape paths | detect geometry changes | shape modification tracking | workbook automation
// Common Searches: Aspose.Cells compare shape paths | detect changes in custom shape geometry .NET | how to get shape path signature Aspose.Cells | track shape geometry modifications in Excel workbook | C# example for shape geometry change detection
// Developer Intent: Identify whether a custom shape's geometry has been altered by comparing its path data before and after an edit.
// Use Cases: Validate that custom diagram shapes remain unchanged after applying workbook transformations. | Implement change‑tracking for engineering schematics by storing and comparing geometry signatures. | Trigger conditional formatting or alerts when a shape's path collection is modified programmatically.
// AI Prompts: Generate a hash for a CustomGeometry object's paths to enable fast change detection in Aspose.Cells. | Show how to serialize ShapePath data to JSON and compare two versions to find geometry differences. | Provide C# code that restores a shape's geometry from a previously saved signature using Aspose.Cells.

using System;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// This example creates a workbook, adds a NotPrimitive autoshape with custom geometry, records a string signature of all ShapePath segments, inserts a new rectangular path, records a second signature, and compares the two signatures to determine whether the shape's geometry has changed. The workbook is then saved.
class ShapePathComparison
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a NotPrimitive autoshape (custom geometry) to the worksheet
        Shape shape = worksheet.Shapes.AddAutoShape(AutoShapeType.NotPrimitive, 0, 0, 0, 0, 200, 200);

        // Cast the shape's geometry to CustomGeometry to access the Paths collection
        CustomGeometry customGeometry = shape.Geometry as CustomGeometry;
        if (customGeometry == null)
        {
            Console.WriteLine("The shape does not have custom geometry.");
            return;
        }

        // Capture the geometry signature before any modification
        string beforeSignature = GetGeometrySignature(customGeometry);

        // Modify the geometry: add a new rectangular path
        int newPathIndex = customGeometry.Paths.Add();
        ShapePath newPath = customGeometry.Paths[newPathIndex];
        newPath.MoveTo(0, 0);
        newPath.LineTo(10000, 0);
        newPath.LineTo(10000, 10000);
        newPath.LineTo(0, 10000);
        newPath.Close();

        // Capture the geometry signature after modification
        string afterSignature = GetGeometrySignature(customGeometry);

        // Compare the signatures to detect changes
        bool geometryChanged = !beforeSignature.Equals(afterSignature);
        Console.WriteLine($"Geometry changed: {geometryChanged}");
        Console.WriteLine($"Before: {beforeSignature}");
        Console.WriteLine($"After : {afterSignature}");

        // Save the workbook
        workbook.Save("ShapePathComparison.xlsx");
    }

    // Generates a simple string representation of all paths and their segments
    static string GetGeometrySignature(CustomGeometry geometry)
    {
        StringBuilder sb = new StringBuilder();
        for (int i = 0; i < geometry.Paths.Count; i++)
        {
            ShapePath path = geometry.Paths[i];
            sb.Append($"Path{i}:");
            foreach (ShapeSegmentPath segment in path.PathSegementList)
            {
                sb.Append($"{segment.Type}-");
                foreach (ShapePathPoint pt in segment.Points)
                {
                    sb.Append($"({pt.X},{pt.Y})");
                }
                sb.Append(";");
            }
            sb.Append("|");
        }
        return sb.ToString();
    }
}
