using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class ApplyPatternToPlaceholder
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Put placeholder text into cell A1
        Cell placeholderCell = sheet.Cells["A1"];
        placeholderCell.PutValue("Placeholder");

        // Add a rectangle shape that covers the cell (adjust size as needed)
        // Parameters: shape type, upper left row, upper left column, offsetX, offsetY, width, height
        Shape shape = sheet.Shapes.AddAutoShape(AutoShapeType.Rectangle, placeholderCell.Row, placeholderCell.Column, 0, 0, 100, 30);
        shape.Text = placeholderCell.StringValue; // display the placeholder text inside the shape

        // Set the fill to a pattern with small dots
        shape.Fill.FillType = FillType.Pattern;
        shape.Fill.PatternFill.Pattern = FillPattern.DottedGrid; // small dotted grid pattern
        shape.Fill.PatternFill.ForegroundColor = Color.Black;   // dot color
        shape.Fill.PatternFill.BackgroundColor = Color.White;   // background color

        // Clear the original cell text (optional, since the shape now shows the text)
        placeholderCell.PutValue(string.Empty);

        // Save the workbook
        workbook.Save("PlaceholderPattern.xlsx");
    }
}