using Aspose.Cells;
using Aspose.Cells.Drawing;

class UngroupAndModifyShapes
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        ShapeCollection shapes = worksheet.Shapes;

        // Add two rectangle shapes to the worksheet
        Shape shape1 = shapes.AddRectangle(0, 0, 0, 0, 100, 50);
        Shape shape2 = shapes.AddRectangle(0, 0, 3, 0, 100, 50);

        // Group the two shapes into a GroupShape
        GroupShape groupShape = shapes.Group(new Shape[] { shape1, shape2 });

        // Ungroup the shapes using the GroupShape.Ungroup method
        groupShape.Ungroup();

        // After ungrouping, the individual shapes are back in the Shapes collection.
        // Modify each shape's position (e.g., move 20 pixels right and 10 pixels down)
        foreach (Shape s in shapes)
        {
            s.Left += 20; // shift horizontally
            s.Top += 10;  // shift vertically
        }

        // Save the workbook with the modified shapes
        workbook.Save("UngroupedModifiedShapes.xlsx");
    }
}