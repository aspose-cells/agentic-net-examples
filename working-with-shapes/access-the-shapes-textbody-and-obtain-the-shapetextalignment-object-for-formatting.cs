using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a rectangle shape to the worksheet
        Shape shape = worksheet.Shapes.AddRectangle(1, 0, 1, 0, 100, 50);
        shape.Text = "Sample Text";

        // Access the ShapeTextAlignment object via the shape's TextBody
        ShapeTextAlignment textAlignment = shape.TextBody.TextAlignment;

        // Example formatting: enable text wrapping and set custom margins
        textAlignment.IsTextWrapped = true;
        textAlignment.TopMarginPt = 12.0;
        textAlignment.BottomMarginPt = 12.0;
        textAlignment.LeftMarginPt = 6.0;
        textAlignment.RightMarginPt = 6.0;

        // Save the workbook
        workbook.Save("ShapeTextAlignmentDemo.xlsx");
    }
}