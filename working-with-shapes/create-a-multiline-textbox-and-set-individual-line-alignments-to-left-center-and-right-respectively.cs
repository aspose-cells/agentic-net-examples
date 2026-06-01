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
        // Parameters: upper left row, upper left column, height, width, lower right row, lower right column
        Shape textBox = worksheet.Shapes.AddTextBox(0, 0, 100, 300, 200, 100);

        // Set multiline text using line breaks
        textBox.Text = "Left aligned line\nCenter aligned line\nRight aligned line";

        // Retrieve the collection of paragraphs (each line is a paragraph)
        TextParagraphCollection paragraphs = textBox.TextBody.TextParagraphs;

        // Set individual alignment for each paragraph
        paragraphs[0].AlignmentType = TextAlignmentType.Left;    // First line left aligned
        paragraphs[1].AlignmentType = TextAlignmentType.Center;  // Second line centered
        paragraphs[2].AlignmentType = TextAlignmentType.Right;   // Third line right aligned

        // Save the workbook
        workbook.Save("MultilineTextBoxAlignment.xlsx");
    }
}