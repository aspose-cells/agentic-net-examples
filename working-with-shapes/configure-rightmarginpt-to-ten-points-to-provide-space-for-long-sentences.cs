using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

class ConfigureRightMarginPt
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a rectangle shape to the worksheet
        // Parameters: upper left row, upper left column, top, left, width, height
        Shape shape = worksheet.Shapes.AddRectangle(1, 0, 1, 100, 200, 50);
        shape.Text = "This is a long sentence that needs extra right margin space to avoid clipping.";

        // Access the shape's text alignment settings
        ShapeTextAlignment textAlignment = shape.TextBody.TextAlignment;

        // Set the right margin of the text frame to 10 points
        textAlignment.RightMarginPt = 10.0;

        // Save the workbook to a file
        workbook.Save("RightMarginPtDemo.xlsx");
    }
}