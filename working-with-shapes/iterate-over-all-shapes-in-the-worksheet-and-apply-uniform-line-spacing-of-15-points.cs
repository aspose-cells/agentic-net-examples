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
        Worksheet sheet = workbook.Worksheets[0];

        // Sample shapes with text to demonstrate the effect
        Shape txtBox1 = sheet.Shapes.AddTextBox(0, 0, 0, 0, 200, 100);
        txtBox1.Text = "First line\nSecond line";

        Shape txtBox2 = sheet.Shapes.AddTextBox(5, 0, 5, 0, 200, 100);
        txtBox2.Text = "Another box\nWith two lines";

        // Iterate over all shapes in the worksheet
        foreach (Shape shape in sheet.Shapes)
        {
            // Process only shapes that contain rich text (have text paragraphs)
            if (shape.IsRichText)
            {
                TextParagraphCollection paragraphs = shape.TextBody.TextParagraphs;
                for (int i = 0; i < paragraphs.Count; i++)
                {
                    TextParagraph paragraph = paragraphs[i];
                    // Set line spacing to 1.5 points
                    paragraph.LineSpaceSizeType = LineSpaceSizeType.Points;
                    paragraph.LineSpace = 1.5;
                }
            }
        }

        // Save the workbook with the updated line spacing
        workbook.Save("ShapesLineSpacing.xlsx");
    }
}