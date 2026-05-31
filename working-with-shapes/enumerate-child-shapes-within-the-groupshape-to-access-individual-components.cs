using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    class EnumerateGroupShapes
    {
        static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }

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

                // Access individual shapes via the indexer
                for (int i = 0; i < childShapes.Length; i++)
                {
                    Shape shape = group[i];
                    Console.WriteLine($"Indexer {i}: Name = {shape.Name}, Text = {shape.Text}");
                }

                // Save the workbook
                string outputPath = "EnumeratedGroupShapes.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during processing: {ex.Message}");
            }
        }
    }
}