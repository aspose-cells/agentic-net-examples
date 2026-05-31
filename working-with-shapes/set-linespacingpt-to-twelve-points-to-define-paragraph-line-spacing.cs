using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

class SetParagraphLineSpacing
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add a text box shape to the worksheet
        Shape textBox = sheet.Shapes.AddTextBox(0, 0, 0, 0, 400, 200);
        // Set multi‑line text (two paragraphs)
        textBox.Text = "First paragraph\nSecond paragraph";

        // Access the second paragraph (index 1)
        TextParagraph paragraph = textBox.TextBody.TextParagraphs[1];

        // Define line spacing in points
        paragraph.LineSpaceSizeType = LineSpaceSizeType.Points; // use points as unit
        paragraph.LineSpace = 12; // twelve points line spacing

        // Save the workbook
        workbook.Save("ParagraphLineSpacing.xlsx");

        // Optional: verify the setting by reloading
        Workbook loaded = new Workbook("ParagraphLineSpacing.xlsx");
        TextParagraph loadedParagraph = loaded.Worksheets[0].Shapes[0].TextBody.TextParagraphs[1];
        Console.WriteLine("LineSpace: " + loadedParagraph.LineSpace);
        Console.WriteLine("LineSpaceSizeType: " + loadedParagraph.LineSpaceSizeType);
    }
}