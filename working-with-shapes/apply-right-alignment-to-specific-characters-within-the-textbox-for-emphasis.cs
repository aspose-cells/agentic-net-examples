// Title: Right‑align specific characters in an Excel textbox with Aspose.Cells for .NET (C#)
// Description: A C# snippet that creates a workbook, adds a textbox shape, finds a target substring, applies bold red styling via Font and StyleFlag, sets the paragraph alignment to right, and writes the result to an .xlsx file.
// Keywords: Aspose.Cells C# textbox alignment | format characters in shape Aspose.Cells | right align substring Excel textbox | StyleFlag text formatting .NET | TextParagraph alignment Aspose.Cells | Excel shape text styling C# | Aspose.Cells example workbook textbox
// Common Searches: How to align part of the text in an Aspose.Cells textbox | Apply bold red style to a substring in an Excel shape using C# | Set paragraph alignment for a textbox shape with Aspose.Cells | Use StyleFlag to format characters inside a textbox in .NET
// Developer Intent: Apply bold red styling to a selected range of characters inside a textbox and make the entire paragraph right‑justified using Aspose.Cells.
// Use Cases: Generating a financial report where the word “Total” inside a textbox is highlighted in red bold and the whole note is right‑aligned for emphasis. | Creating an invoice template that emphasizes the amount field within a textbox while keeping the surrounding text left‑aligned. | Designing a dashboard sheet where status messages are partially highlighted and the message block is aligned to the right edge of the sheet.
// AI Prompts: Write C# code with Aspose.Cells that formats a specific substring in a textbox (bold, red) and then sets the textbox paragraph to right alignment. | Show how to use StyleFlag and TextParagraph objects to style and right‑justify selected characters inside an Excel shape using Aspose.Cells for .NET. | Explain the steps to locate a word in a textbox, apply custom font properties, and adjust paragraph alignment for that range with Aspose.Cells.

using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

// A C# snippet that creates a workbook, adds a textbox shape, finds a target substring, applies bold red styling via Font and StyleFlag, sets the paragraph alignment to right, and writes the result to an .xlsx file.
class RightAlignSpecificChars
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a text box shape to the worksheet
        Shape textBox = worksheet.Shapes.AddTextBox(0, 0, 100, 100, 300, 100);
        textBox.Text = "Important: Align this part right";

        // Determine the start index and length of the characters to emphasize
        int startIndex = textBox.Text.IndexOf("Align");
        int length = "Align this part".Length;

        // Prepare a font with the desired emphasis (bold and red)
        Aspose.Cells.Font emphasisFont = textBox.Font;
        emphasisFont.IsBold = true;
        emphasisFont.Color = Color.Red;

        // Specify which font properties should be applied
        StyleFlag flag = new StyleFlag();
        flag.FontBold = true;
        flag.FontColor = true;

        // Apply the formatting to the selected characters
        textBox.FormatCharacters(startIndex, length, emphasisFont, flag);

        // Set the paragraph alignment of the whole text box to right
        TextParagraph paragraph = textBox.TextBody.TextParagraphs[0];
        paragraph.AlignmentType = TextAlignmentType.Right;

        // Save the workbook
        workbook.Save("RightAlignSpecificChars.xlsx");
    }
}
