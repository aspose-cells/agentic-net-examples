// Title: How to center text horizontally and vertically in a rectangle shape using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set shape text alignment to center both horizontally and vertically | center text inside a rectangle shape in an Excel worksheet with Aspose.Cells | C# code to align shape text to middle and center using Aspose.Cells API | Aspose.Cells shape text alignment properties example for .NET
// Tags: Aspose.Cells shape text horizontal alignment property | Aspose.Cells shape text vertical alignment property | C# set shape text alignment center | Excel rectangle shape text formatting Aspose.Cells | programmatic shape text alignment .NET

using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, adds a rectangle shape to the first worksheet, assigns "Centered Text" to the shape, sets TextHorizontalAlignment and TextVerticalAlignment to Center, and saves the workbook as AlignedShape.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle rule: create)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Add a rectangle shape to the worksheet (example shape)
        // Parameters: shape type, upper left row, upper left column, upper left offset X, upper left offset Y, width, height
        Shape shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 2, 2, 0, 0, 150, 80);
        shape.Text = "Centered Text";

        // Set horizontal text alignment to center
        shape.TextHorizontalAlignment = TextAlignmentType.Center;

        // Set vertical text alignment to middle (center)
        shape.TextVerticalAlignment = TextAlignmentType.Center;

        // Save the workbook (lifecycle rule: save)
        workbook.Save("AlignedShape.xlsx");
    }
}
