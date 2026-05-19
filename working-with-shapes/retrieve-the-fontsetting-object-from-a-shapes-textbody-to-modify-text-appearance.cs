using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

class RetrieveFontSettingFromShape
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a text box shape to the worksheet
        // Parameters: upper left row, upper left column, upper left offset in pixels,
        // lower right row, lower right column, lower right offset in pixels
        Shape shape = worksheet.Shapes.AddTextBox(0, 0, 0, 100, 200, 0);
        shape.Text = "Aspose Cells FontSetting Demo";

        // Retrieve the FontSettingCollection from the shape's TextBody
        FontSettingCollection textBody = shape.TextBody;

        // Optionally set the whole text (not required if shape.Text already set)
        // textBody.Text = shape.Text;

        // Retrieve a specific FontSetting object by index (e.g., first character)
        // The collection contains a FontSetting for each character range.
        FontSetting fontSetting = textBody[0];

        // Modify the font appearance of the selected FontSetting
        fontSetting.Font.Name = "Calibri";
        fontSetting.Font.Size = 14;
        fontSetting.Font.Color = Color.Blue;
        fontSetting.Font.IsBold = true;

        // Save the workbook
        workbook.Save("ShapeFontSettingDemo.xlsx");
    }
}