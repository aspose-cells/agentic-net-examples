using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsGroupingDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Get the shapes collection of the worksheet
            ShapeCollection shapes = sheet.Shapes;

            // Add a textbox (cast to TextBox for textbox‑specific properties if needed)
            TextBox textBox = (TextBox)shapes.AddTextBox(2, 0, 2, 0, 100, 200);
            textBox.Text = "Grouped TextBox";

            // Add other shapes to be grouped with the textbox
            Shape rectangle = shapes.AddRectangle(5, 0, 5, 0, 80, 60);
            Shape oval = shapes.AddOval(8, 0, 8, 0, 80, 60);

            // Create an array containing the shapes to group
            Shape[] groupItems = new Shape[] { textBox, rectangle, oval };

            // Group the shapes together
            GroupShape groupShape = shapes.Group(groupItems);

            // Optional: set properties on the group container
            groupShape.Name = "MyShapeGroup";
            groupShape.Placement = PlacementType.FreeFloating;

            // Save the workbook
            workbook.Save("GroupedShapes.xlsx");
        }
    }
}