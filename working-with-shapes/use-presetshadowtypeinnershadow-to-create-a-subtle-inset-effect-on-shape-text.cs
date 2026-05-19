using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add a rectangle shape to the worksheet
        // Parameters: upperLeftRow, top, upperLeftColumn, left, height, width
        Shape shape = sheet.Shapes.AddRectangle(2, 2, 2, 2, 200, 100);

        // Set the shape's text
        shape.TextBody.Text = "Inset Shadow";

        // Get the TextOptions of the first paragraph in the shape's text body
        TextOptions textOptions = shape.TextBody[0].TextOptions;

        // Configure basic font properties
        textOptions.Name = "Calibri";
        textOptions.Size = 24;
        textOptions.IsBold = true;
        textOptions.Color = Color.White;

        // Apply an inner (inset) shadow preset to the text
        textOptions.Shadow.PresetType = PresetShadowType.InsideCenter;

        // Optional: make the inner shadow subtle
        textOptions.Shadow.Transparency = 0.3; // 30% transparent
        textOptions.Shadow.Size = 1.0;         // size is ignored for inner shadows but set a default

        // Save the workbook
        workbook.Save("InnerShadowShapeText.xlsx", SaveFormat.Xlsx);
    }
}