using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Get the shapes collection of the worksheet
        ShapeCollection shapes = sheet.Shapes;

        // Add three shapes to be grouped
        // Rectangle
        Shape rect = shapes.AddRectangle(2, 0, 2, 0, 60, 80);
        // Oval
        Shape oval = shapes.AddOval(6, 0, 2, 0, 60, 80);
        // Line
        Shape line = shapes.AddLine(10, 0, 2, 0, 0, 120);

        // Group the three shapes using ShapeCollection.Group
        GroupShape group = shapes.Group(new Shape[] { rect, oval, line });

        // Assign a descriptive name to the group for later reference
        group.Name = "MyCompositeGroup";

        // Save the workbook
        workbook.Save("GroupedShapes.xlsx");
    }
}