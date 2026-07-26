// Title: Set 1.5 line spacing for shape text in Excel with Aspose.Cells for .NET
// Description: Creates a workbook, adds a textbox shape, fills it with multi‑line text, and applies 150 % (1.5 lines) line spacing to each paragraph using TextOptions.LineSpacing (or TextParagraph properties). The workbook is then saved as an .xlsx file.
// Keywords: Aspose.Cells .NET line spacing | shape text formatting | textbox paragraph spacing | TextOptions.LineSpacing | Excel shape text height | C# Aspose.Cells example
// Common Searches: Aspose.Cells set line spacing in shape | C# textbox line height 1.5 lines | How to change paragraph spacing in Excel shape using Aspose | TextOptions.LineSpacing example Aspose.Cells | Adjust shape text line spacing programmatically
// Developer Intent: Apply 150 % line spacing to all paragraphs inside a shape's text.
// Use Cases: Designing reports where textbox notes need consistent 1.5‑line spacing for readability. | Building dashboards with shape captions that require uniform line height. | Generating templates where multi‑line shape text must follow a specific spacing rule.
// AI Prompts: Provide C# code that sets 1.5 line spacing for a textbox shape using Aspose.Cells TextOptions.LineSpacing. | Explain the difference between TextOptions.LineSpacing and TextParagraph.LineSpace for shape text formatting. | Show how to format shape text with 150 % line spacing in an Excel workbook via Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

// Creates a workbook, adds a textbox shape, fills it with multi‑line text, and applies 150 % (1.5 lines) line spacing to each paragraph using TextOptions.LineSpacing (or TextParagraph properties). The workbook is then saved as an .xlsx file.
class AdjustLineSpacing
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a text box shape to the worksheet
        Shape textBox = worksheet.Shapes.AddTextBox(0, 0, 0, 0, 300, 200);
        textBox.Text = "First line\nSecond line\nThird line";

        // Set line spacing to 1.5 lines (150% of the font size) for each paragraph
        foreach (TextParagraph paragraph in textBox.TextBody.TextParagraphs)
        {
            paragraph.LineSpaceSizeType = LineSpaceSizeType.Percentage; // Use percentage unit
            paragraph.LineSpace = 150; // 150% = 1.5 lines
        }

        // Save the workbook
        workbook.Save("LineSpacingDemo.xlsx");
    }
}
