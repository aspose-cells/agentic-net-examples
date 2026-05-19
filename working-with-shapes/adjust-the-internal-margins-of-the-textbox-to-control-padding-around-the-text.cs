using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

class AdjustTextboxMargins
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add a textbox shape to the worksheet
        // Parameters: upper left row, upper left column, upper left row offset, upper left column offset, width, height
        TextBox textBox = sheet.Shapes.AddTextBox(1, 1, 2, 2, 300, 150);
        textBox.Text = "This is a sample text inside the textbox.\nSecond line of the paragraph.";

        // Adjust internal margins (padding) of the textbox in points
        textBox.TextBoxOptions.LeftMarginPt = 15;   // left padding
        textBox.TextBoxOptions.RightMarginPt = 15;  // right padding
        textBox.TextBoxOptions.TopMarginPt = 10;    // top padding
        textBox.TextBoxOptions.BottomMarginPt = 10; // bottom padding

        // Optionally adjust paragraph-specific margins for finer control
        TextParagraph paragraph = textBox.TextBody.TextParagraphs[0];
        paragraph.LeftMargin = 5;   // additional left indent within the paragraph (points)
        paragraph.RightMargin = 5;  // additional right indent within the paragraph (points)

        // Save the workbook with the adjusted textbox margins
        workbook.Save("AdjustedTextboxMargins.xlsx");
    }
}