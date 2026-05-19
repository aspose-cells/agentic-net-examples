using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

class ApplyCharacterSpacing
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add a text box shape
        Shape shape1 = sheet.Shapes.AddTextBox(0, 0, 0, 0, 200, 100);
        shape1.Text = "First shape text";

        // Add a rectangle shape with text
        Shape shape2 = sheet.Shapes.AddRectangle(1, 0, 0, 150, 200, 0);
        shape2.Text = "Second shape text";

        // Apply character spacing of 1.2 points to all shapes that contain text
        foreach (Shape shp in sheet.Shapes)
        {
            if (!string.IsNullOrEmpty(shp.Text))
            {
                shp.TextOptions.Spacing = 1.2;
            }
        }

        // Save the workbook
        workbook.Save("ShapesWithSpacing.xlsx");
    }
}