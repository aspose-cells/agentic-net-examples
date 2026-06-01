using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a textbox shape to the worksheet
        // Parameters: topRow, top (pixel offset), leftColumn, left (pixel offset), height (pixel), width (pixel)
        TextBox textBox = worksheet.Shapes.AddTextBox(1, 0, 1, 0, 100, 200);

        // Set the text of the textbox using the TextBody.Text property
        textBox.TextBody.Text = "This text is set via TextBody.Text property.";

        // Save the workbook to a file
        workbook.Save("TextBoxWithTextBody.xlsx");
    }
}