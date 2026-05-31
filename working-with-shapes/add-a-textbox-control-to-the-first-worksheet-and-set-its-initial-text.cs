using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a TextBox shape to the worksheet
        // Parameters: topRow, top (pixel offset), leftColumn, left (pixel offset), height, width
        TextBox textBox = worksheet.Shapes.AddTextBox(1, 0, 1, 0, 100, 200);

        // Set the initial text of the TextBox
        textBox.Text = "Initial Text";

        // Save the workbook
        workbook.Save("TextBoxDemo.xlsx");
    }
}