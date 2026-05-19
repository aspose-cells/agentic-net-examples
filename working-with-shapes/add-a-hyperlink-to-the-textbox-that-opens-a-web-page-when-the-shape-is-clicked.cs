using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsHyperlinkExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a TextBox shape to the worksheet
            // Parameters: upper left row, upper left column, top offset, left offset, width, height
            TextBox textBox = sheet.Shapes.AddTextBox(2, 2, 10, 10, 200, 50);
            textBox.Text = "Click here to visit Aspose";

            // Add a hyperlink to the TextBox that opens a web page when clicked
            textBox.AddHyperlink("https://www.aspose.com/");

            // Save the workbook to a file
            workbook.Save("TextBoxWithHyperlink.xlsx");
        }
    }
}