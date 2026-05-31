using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

class TextBoxWrapDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add a TextBox shape to the worksheet
        // Parameters: upperRow, leftColumn, upperRowOffset, leftColumnOffset, height, width
        TextBox textBox = sheet.Shapes.AddTextBox(1, 0, 0, 0, 100, 200);

        // Set a long text that needs wrapping
        textBox.Text = "This is a very long sentence that should automatically wrap inside the textbox shape when text wrapping is enabled.";

        // Enable text wrapping within the shape
        textBox.TextBoxOptions.WrapTextInShape = true;

        // Prevent text from overflowing the shape boundaries
        textBox.TextBoxOptions.AllowTextToOverflow = false;

        // Save the workbook
        workbook.Save("TextBoxWrapDemo.xlsx");
    }
}