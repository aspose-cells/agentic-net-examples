using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class SetMultilingualFonts
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a rectangle shape to the worksheet
        Shape shape = worksheet.Shapes.AddRectangle(1, 0, 1, 0, 200, 100);
        shape.Text = "English and Japanese text";

        // Specify the FarEast (CJK) font name
        shape.TextOptions.FarEastName = "MS Mincho";

        // Specify the Latin (Western) font name
        shape.TextOptions.LatinName = "Arial";

        // Save the workbook to a file
        workbook.Save("MultilingualFonts.xlsx");
    }
}