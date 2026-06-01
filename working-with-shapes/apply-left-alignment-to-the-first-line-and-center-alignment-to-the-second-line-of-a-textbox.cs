using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

class TextBoxAlignmentExample
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a textbox shape to the worksheet
        int textboxIndex = worksheet.TextBoxes.Add(1, 1, 200, 100);
        TextBox textbox = worksheet.TextBoxes[textboxIndex];

        // Set the textbox text with two lines
        textbox.Text = "First line\nSecond line";

        // Retrieve the paragraphs (each line is a separate paragraph)
        TextParagraphCollection paragraphs = textbox.TextBody.TextParagraphs;

        // Apply left alignment to the first line (paragraph)
        paragraphs[0].AlignmentType = TextAlignmentType.Left;

        // Apply center alignment to the second line (paragraph)
        paragraphs[1].AlignmentType = TextAlignmentType.Center;

        // Save the workbook
        workbook.Save("AlignedTextBox.xlsx");
    }
}