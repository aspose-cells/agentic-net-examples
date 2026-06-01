using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet (Sheet1)
        Worksheet sheet = workbook.Worksheets[0];

        // Add a TextBox anchored to cell B2 (row index 1, column index 1)
        // Parameters: topRow, top (pixel offset), leftColumn, left (pixel offset), height (pixel), width (pixel)
        TextBox textBox = sheet.Shapes.AddTextBox(1, 0, 1, 0, 100, 50);

        // Set the width of the TextBox to 200 points
        textBox.WidthPt = 200;

        // Optional: set some sample text
        textBox.Text = "Sample TextBox";

        // Save the workbook
        workbook.Save("TextBoxDemo.xlsx");
    }
}