// Title: Enumerate Child Shapes in a GroupShape with Aspose.Cells for .NET
// Description: Creates a workbook, adds a rectangle and an oval, groups them into a GroupShape, then lists each child shape using GetGroupedShapes() and the GroupShape indexer, displaying type and AlternativeText before saving the file.
// Keywords: Aspose.Cells | GroupShape | GetGroupedShapes | .NET | C# | enumerate shapes | Excel shape grouping | access child shapes | shape iteration | Aspose.Cells API
// Common Searches: How to get individual shapes from a GroupShape using Aspose.Cells | Aspose.Cells C# enumerate shapes in a grouped object | Retrieve child shapes of a grouped shape in Excel with Aspose | List shapes inside a GroupShape Aspose.Cells .NET | Iterate over grouped shapes in a worksheet using Aspose.Cells
// Developer Intent: Retrieve and manipulate each shape contained in a GroupShape.
// Use Cases: Read type and AlternativeText of all grouped shapes for reporting | Change properties (fill color, size) of specific child shapes after enumeration | Delete a shape from a group based on its AlternativeText | Export shape metadata to CSV or JSON for external analysis | Apply conditional formatting to grouped shapes programmatically
// AI Prompts: Generate C# code with Aspose.Cells that enumerates child shapes of a GroupShape and sets each shape's fill color to blue. | Provide an example that iterates through grouped shapes and removes the shape whose AlternativeText equals "Oval1" using Aspose.Cells. | Create code that extracts the Type and AlternativeText of every shape in a GroupShape and writes the data to a CSV file.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a workbook, adds a rectangle and an oval, groups them into a GroupShape, then lists each child shape using GetGroupedShapes() and the GroupShape indexer, displaying type and AlternativeText before saving the file.
class EnumerateGroupShapes
{
    public static void Run()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add two shapes to the worksheet
            Shape rect = sheet.Shapes.AddRectangle(2, 0, 2, 0, 80, 60);
            rect.AlternativeText = "Rectangle1";

            Shape oval = sheet.Shapes.AddOval(4, 0, 4, 0, 80, 60);
            oval.AlternativeText = "Oval1";

            // Group the shapes into a GroupShape
            GroupShape group = sheet.Shapes.Group(new Shape[] { rect, oval });

            // Enumerate child shapes using GetGroupedShapes()
            Shape[] groupedShapes = group.GetGroupedShapes();
            Console.WriteLine($"Group contains {groupedShapes.Length} shapes (GetGroupedShapes):");
            foreach (Shape shape in groupedShapes)
            {
                Console.WriteLine($"Type: {shape.Type}, AltText: {shape.AlternativeText}");
            }

            // Enumerate child shapes using the indexer
            Console.WriteLine("Group enumeration using indexer:");
            for (int i = 0; i < groupedShapes.Length; i++)
            {
                Shape shape = group[i];
                Console.WriteLine($"Index {i}: Type: {shape.Type}, AltText: {shape.AlternativeText}");
            }

            // Save the workbook
            workbook.Save("EnumeratedGroupShapes.xlsx", SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

class Program
{
    static void Main(string[] args)
    {
        EnumerateGroupShapes.Run();
    }
}
