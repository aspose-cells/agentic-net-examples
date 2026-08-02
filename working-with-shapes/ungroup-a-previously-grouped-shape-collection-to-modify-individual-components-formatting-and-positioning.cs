using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapeUngroupDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Load a template workbook if it exists; otherwise create a new one
                string templatePath = "template.xlsx";
                Workbook workbook = File.Exists(templatePath) ? new Workbook(templatePath) : new Workbook();

                Worksheet worksheet = workbook.Worksheets[0];

                // Add two shapes to the worksheet
                Shape rect = worksheet.Shapes.AddRectangle(2, 0, 2, 0, 100, 60);
                rect.Name = "RectangleShape";

                Shape oval = worksheet.Shapes.AddOval(6, 0, 2, 0, 80, 80);
                oval.Name = "OvalShape";

                // Group the two shapes
                GroupShape group = worksheet.Shapes.Group(new Shape[] { rect, oval });
                group.Name = "MyGroup";

                // OPTIONAL: Retrieve the shapes that are part of the group before ungrouping
                Shape[] groupedShapes = group.GetGroupedShapes();

                // Save the workbook to verify the result
                string outputPath = "output.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}