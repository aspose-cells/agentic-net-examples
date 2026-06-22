using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

class RetrieveShapeMargins
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a rectangle shape to the worksheet
        // Parameters: upper left row, upper left column, width (pixels), height (pixels), upper left row offset, upper left column offset
        Shape shape = worksheet.Shapes.AddRectangle(1, 1, 100, 80, 0, 0);
        shape.Text = "Sample text";

        // Access the text alignment object which holds the margin properties
        ShapeTextAlignment alignment = shape.TextBody.TextAlignment;

        // Retrieve the current margin values (in points)
        double topMarginPt    = alignment.TopMarginPt;
        double leftMarginPt   = alignment.LeftMarginPt;
        double bottomMarginPt = alignment.BottomMarginPt;
        double rightMarginPt  = alignment.RightMarginPt;

        // Display the retrieved margin values
        Console.WriteLine($"TopMarginPt:    {topMarginPt}");
        Console.WriteLine($"LeftMarginPt:   {leftMarginPt}");
        Console.WriteLine($"BottomMarginPt: {bottomMarginPt}");
        Console.WriteLine($"RightMarginPt:  {rightMarginPt}");

        // Save the workbook (optional, just to complete the lifecycle)
        workbook.Save("ShapeMarginsDemo.xlsx");
    }
}