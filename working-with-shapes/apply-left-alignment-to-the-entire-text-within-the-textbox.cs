using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

class LeftAlignTextboxDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a text box shape to the worksheet
        // Parameters: upper left row, upper left column, top, left, width, height (in points)
        Shape textBox = worksheet.Shapes.AddTextBox(0, 0, 100, 100, 200, 100);

        // Set the text with multiple lines (each line is a separate paragraph)
        textBox.Text = "First line\nSecond line\nThird line";

        // Apply left alignment to every paragraph in the text box
        foreach (TextParagraph paragraph in textBox.TextBody.TextParagraphs)
        {
            paragraph.AlignmentType = TextAlignmentType.Left;
        }

        // Save the workbook
        workbook.Save("LeftAlignedTextbox.xlsx");
    }
}