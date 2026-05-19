using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

class SetCharacterSpacing
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a rectangle shape that will contain text
        Shape shape = worksheet.Shapes.AddRectangle(1, 0, 0, 100, 200, 0);
        shape.Text = "Sample Text";

        // Access the shape's TextOptions and set character spacing to 5 points
        TextOptions textOptions = shape.TextOptions;
        textOptions.Spacing = 5.0; // 5 points spacing between characters

        // Save the workbook
        workbook.Save("CharacterSpacingDemo.xlsx");
    }
}