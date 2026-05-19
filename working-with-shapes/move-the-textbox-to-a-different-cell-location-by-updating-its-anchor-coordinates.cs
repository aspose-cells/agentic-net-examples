using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class MoveTextBoxDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a textbox at an initial location (row 2, column 2)
        // Parameters: topRow, top (pixel offset), leftColumn, left (pixel offset), height, width
        TextBox textBox = worksheet.Shapes.AddTextBox(2, 0, 2, 0, 100, 200);
        textBox.Text = "Original Position";

        // Move the textbox to a new cell location (row 10, column 5)
        // Row and column indices are zero‑based, so row 10 => index 9, column 5 => index 4
        textBox.UpperLeftRow = 9;
        textBox.UpperLeftColumn = 4;

        // (Optional) Adjust size after moving
        textBox.Height = 120;
        textBox.Width = 250;

        // Save the workbook with the moved textbox
        workbook.Save("MovedTextBox.xlsx");
    }
}