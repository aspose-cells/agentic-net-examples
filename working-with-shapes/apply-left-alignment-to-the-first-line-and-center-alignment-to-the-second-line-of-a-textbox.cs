// Title: Set left alignment for the first line and center alignment for the second line of a TextBox in Aspose.Cells for .NET (C#)
// Description: Creates a workbook, adds a TextBox shape with two lines, retrieves the TextParagraphCollection, applies left alignment to the first paragraph and center alignment to the second, then saves the workbook as an XLSX file.
// Keywords: Aspose.Cells | C# | TextBox | paragraph alignment | TextAlignmentType | left alignment | center alignment | shape text formatting | Aspose.Cells .NET | TextParagraphCollection
// Common Searches: Aspose.Cells align first line left second line center | C# set paragraph alignment in TextBox Aspose.Cells | How to change text alignment inside a shape using Aspose.Cells | TextBox paragraph alignment .NET | Apply different alignments to lines in Aspose.Cells textbox
// Developer Intent: Apply distinct horizontal alignments to individual paragraphs within a TextBox shape using Aspose.Cells for .NET.
// Use Cases: Generate a report where the title line is left‑aligned and the subtitle is centered inside a textbox for visual hierarchy. | Create a dashboard sheet with a textbox that contains a header left‑aligned and a description centered to improve readability. | Automate spreadsheet templates that require different alignments for multiple lines within a shape.
// AI Prompts: Show C# code that sets left alignment for the first paragraph and center alignment for the second paragraph of a TextBox using Aspose.Cells. | Explain how to access a TextBox's TextParagraphCollection and modify the AlignmentType of specific lines in Aspose.Cells for .NET. | Provide an example of applying different horizontal alignments to multiple lines inside a textbox shape in a workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

// Creates a workbook, adds a TextBox shape with two lines, retrieves the TextParagraphCollection, applies left alignment to the first paragraph and center alignment to the second, then saves the workbook as an XLSX file.
class TextBoxParagraphAlignmentDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a text box shape
        // Parameters: upper left row, upper left column, top, left, height, width
        Shape textBox = worksheet.Shapes.AddTextBox(0, 0, 100, 300, 200, 100);

        // Set the text with two lines
        textBox.Text = "First line\nSecond line";

        // Get the collection of paragraphs inside the text box
        TextParagraphCollection paragraphs = textBox.TextBody.TextParagraphs;

        // Apply left alignment to the first line (paragraph)
        paragraphs[0].AlignmentType = TextAlignmentType.Left;

        // Apply center alignment to the second line (paragraph)
        if (paragraphs.Count > 1)
        {
            paragraphs[1].AlignmentType = TextAlignmentType.Center;
        }

        // Save the workbook
        workbook.Save("TextBoxParagraphAlignmentDemo.xlsx");
    }
}
