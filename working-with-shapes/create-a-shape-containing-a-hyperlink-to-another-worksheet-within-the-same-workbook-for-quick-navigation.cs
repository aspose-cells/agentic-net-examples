using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet and add a second worksheet for navigation
        Worksheet sheet1 = workbook.Worksheets[0];
        Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");

        // Add a rectangle shape to the first worksheet
        // Parameters: upper left row, upper left column, width, height, upper left offset X, upper left offset Y
        Shape shape = sheet1.Shapes.AddRectangle(2, 2, 120, 40, 0, 0);
        shape.Text = "Go to Sheet2";

        // Add a hyperlink to the shape that points to cell A1 of the second worksheet
        Hyperlink hyperlink = shape.AddHyperlink("Sheet2!A1");
        hyperlink.ScreenTip = "Navigate to Sheet2";

        // Save the workbook
        workbook.Save("ShapeHyperlinkToWorksheet.xlsx");
    }
}