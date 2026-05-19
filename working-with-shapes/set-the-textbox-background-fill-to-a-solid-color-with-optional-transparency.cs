using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class SetTextboxBackground
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add a textbox shape to the worksheet
        // Parameters: drawing type, upper left row, upper left column, top, left, width, height
        Shape textbox = sheet.Shapes.AddShape(MsoDrawingType.TextBox, 2, 0, 2, 0, 200, 100);

        // Configure the shape to use a solid fill
        textbox.Fill.FillType = FillType.Solid;

        // Set the solid fill color (example: a custom ARGB color)
        textbox.Fill.SolidFill.Color = Color.FromArgb(255, 100, 150, 200);

        // Optional: set transparency (0.0 = opaque, 1.0 = fully transparent)
        textbox.Fill.SolidFill.Transparency = 0.3; // 30% transparent

        // Add some text to the textbox (optional)
        textbox.Text = "Sample TextBox";

        // Save the workbook to verify the textbox appearance
        workbook.Save("TextboxBackgroundSolid.xlsx");
    }
}