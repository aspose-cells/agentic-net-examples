using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;
using System.Drawing;

class ApplyOuterShadowToShapeText
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add a rectangle shape to the worksheet
        // Parameters: upperLeftRow, top, upperLeftColumn, left, height, width
        Shape shape = sheet.Shapes.AddAutoShape(AutoShapeType.Rectangle, 4, 4, 4, 4, 100, 700);
        shape.Fill.FillType = FillType.None; // No fill for clarity
        shape.Text = "Hello World !!!";

        // Get the FontSetting for the shape's text
        FontSetting fontSetting = shape.Characters(0, shape.Text.Length);
        // Access TextOptions from the FontSetting
        TextOptions textOptions = fontSetting.TextOptions;

        // Set desired text formatting (optional)
        textOptions.Name = "Calibri";
        textOptions.Size = 54;
        textOptions.IsBold = true;
        textOptions.Color = Color.Green;

        // Apply an outer preset shadow to the text
        // Choose any outer shadow type, e.g., OffsetBottom
        textOptions.Shadow.PresetType = PresetShadowType.OffsetBottom;

        // Save the workbook
        workbook.Save("ShapeTextWithOuterShadow.xlsx", SaveFormat.Xlsx);
    }
}