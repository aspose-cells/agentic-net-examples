using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

class AdjustTopMarginDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add a rectangle shape to the worksheet
        Shape shape = sheet.Shapes.AddRectangle(1, 1, 100, 100, 0, 0);
        shape.Text = "Sample Text";

        // Adjust the top margin of the shape's text to 5 points
        shape.TextBody.TextAlignment.TopMarginPt = 5.0;

        // Save the workbook with the updated top margin
        workbook.Save("AdjustedTopMargin.xlsx");
    }
}