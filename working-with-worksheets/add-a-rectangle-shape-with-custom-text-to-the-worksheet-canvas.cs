using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a rectangle shape (row index 2, column index 2) with height 100px and width 150px
        RectangleShape rectangle = worksheet.Shapes.AddRectangle(2, 0, 2, 0, 100, 150);

        // Set custom text for the rectangle
        rectangle.Text = "Custom rectangle text";

        // Optional: style the shape
        rectangle.Fill.FillType = FillType.Solid;
        rectangle.Fill.SolidFill.Color = System.Drawing.Color.LightYellow;
        rectangle.Line.DashStyle = MsoLineDashStyle.Solid;
        rectangle.Line.Weight = 1.5;

        // Save the workbook
        workbook.Save("RectangleWithText.xlsx");
    }
}