using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

class AutoFitTextboxDemo
{
    public static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a text box shape to the worksheet
        // Parameters: upper left row, upper left column, top offset, left offset, width, height
        Shape textBox = worksheet.Shapes.AddTextBox(1, 1, 100, 100, 200, 50);

        // Set the text that will determine the required size
        textBox.Text = "This is a long text that should cause the textbox to automatically resize based on its content length.";

        // Enable automatic resizing to fit the text
        textBox.TextBoxOptions.ResizeToFitText = true;

        // Recalculate the shape size to fit the current text
        textBox.FitToTextSize();

        // Save the workbook
        workbook.Save("AutoFitTextboxDemo.xlsx");
    }
}