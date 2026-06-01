using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

class RightAlignSpecificCharacters
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a text box shape to the worksheet
        // Parameters: upper left row, upper left column, upper left offset X, upper left offset Y, width, height
        Shape textBox = worksheet.Shapes.AddTextBox(0, 0, 100, 100, 300, 100);
        textBox.Text = "Important: Review the report";

        // Align the whole paragraph to the right
        TextParagraph paragraph = textBox.TextBody.TextParagraphs[0];
        paragraph.AlignmentType = TextAlignmentType.Right;

        // Prepare a font with the desired emphasis style
        Aspose.Cells.Font emphasisFont = textBox.Font;
        emphasisFont.IsBold = true;          // Make text bold
        emphasisFont.Color = Color.Red;      // Change color to red
        emphasisFont.Size = 12;              // Set font size (optional)

        // Define which characters to format (e.g., the word "Important")
        int startIndex = 0;                                   // start at the first character
        int length = "Important".Length;                      // length of the word to emphasize

        // Specify which font attributes should be applied
        StyleFlag flag = new StyleFlag();
        flag.FontBold = true;
        flag.FontColor = true;
        flag.FontSize = true;
        flag.FontName = true;

        // Apply the formatting to the selected characters
        textBox.FormatCharacters(startIndex, length, emphasisFont, flag);

        // Save the workbook
        workbook.Save("RightAlignedEmphasis.xlsx");
    }
}