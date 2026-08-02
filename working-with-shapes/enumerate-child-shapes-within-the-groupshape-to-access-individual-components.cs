// Title: Enumerate Child Shapes in a GroupShape using Aspose.Cells for .NET
// Description: Shows how to create a workbook, add rectangle and oval shapes, group them into a GroupShape, enumerate the grouped shapes with GetGroupedShapes() and the GroupShape indexer, display each shape’s type, alternative text, name, and text, and save the result as EnumerateGroupShapes.xlsx.
// Keywords: Aspose.Cells | C# | .NET | GroupShape | GetGroupedShapes | enumerate shapes | shape grouping | Excel shape iteration | shape indexer | workbook saving
// Common Searches: list shapes inside a GroupShape Aspose.Cells | how to enumerate child shapes of a GroupShape in C# | Aspose.Cells GetGroupedShapes example | iterate over grouped shapes in Excel using .NET | access individual shapes after grouping with Aspose.Cells
// Developer Intent: The developer wants to retrieve and work with each individual shape that belongs to a GroupShape.
// Use Cases: Display properties (type, alt text, name) of all shapes within a grouped object for debugging. | Apply formatting or data to specific child shapes after they have been grouped. | Generate a report of grouped shape composition before exporting the workbook.
// AI Prompts: Write C# code that adds three shapes, groups them, and changes the fill color of the second child shape using the GroupShape indexer. | Provide an example that ungroups a GroupShape and accesses the original shapes with Aspose.Cells. | Explain how to safely enumerate child shapes when a GroupShape may contain zero elements.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // Shows how to create a workbook, add rectangle and oval shapes, group them into a GroupShape, enumerate the grouped shapes with GetGroupedShapes() and the GroupShape indexer, display each shape’s type, alternative text, name, and text, and save the result as EnumerateGroupShapes.xlsx.
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
                Shape[] childShapes = group.GetGroupedShapes();
                Console.WriteLine($"Group contains {childShapes.Length} shapes:");
                foreach (Shape child in childShapes)
                {
                    Console.WriteLine($"Type: {child.Type}, AltText: {child.AlternativeText}");
                }

                // Access child shapes via the indexer
                for (int i = 0; i < childShapes.Length; i++)
                {
                    Shape s = group[i];
                    Console.WriteLine($"Indexer [{i}] Name: {s.Name}, Text: {s.Text}");
                }

                // Save the workbook
                workbook.Save("EnumerateGroupShapes.xlsx", SaveFormat.Xlsx);
                Console.WriteLine("Workbook saved as EnumerateGroupShapes.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Entry point required for the application
        static void Main(string[] args)
        {
            Run();
        }
    }
}
