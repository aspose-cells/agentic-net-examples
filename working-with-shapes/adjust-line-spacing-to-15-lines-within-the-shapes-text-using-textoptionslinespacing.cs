using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

class AdjustLineSpacingDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a text box shape to the worksheet
        // Parameters: upper left row, upper left column, top, left, width, height
        Shape shape = worksheet.Shapes.AddTextBox(0, 0, 0, 0, 300, 150);
        shape.Text = "First line\nSecond line\nThird line";

        // Access all paragraphs within the shape's text
        TextParagraphCollection paragraphs = shape.TextBody.TextParagraphs;

        // Set line spacing to 1.5 lines (150%) for each paragraph
        foreach (TextParagraph paragraph in paragraphs)
        {
            // Use Percentage unit type and set value to 150 (i.e., 150% of font size)
            paragraph.LineSpaceSizeType = LineSpaceSizeType.Percentage;
            paragraph.LineSpace = 150;
        }

        // Save the workbook
        workbook.Save("AdjustedLineSpacing.xlsx");
    }
}