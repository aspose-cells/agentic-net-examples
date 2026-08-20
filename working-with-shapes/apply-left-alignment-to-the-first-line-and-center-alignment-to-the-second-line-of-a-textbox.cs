// Title: Align TextBox Paragraphs in Aspose.Cells for .NET – Left‑align First Line, Center‑align Second Line
// Description: Creates a workbook, adds a TextBox shape, inserts two lines of text, and sets the first paragraph to left alignment and the second paragraph to center alignment using Aspose.Cells for .NET before saving the file.
// Keywords: Aspose.Cells TextBox alignment | C# paragraph alignment Aspose.Cells | left align first line Aspose.Cells | center align second line Aspose.Cells | multiline textbox shape .NET | Excel shape text formatting | Aspose.Cells example C#
// Common Searches: Aspose.Cells set left alignment for first line of textbox | center align second paragraph in Aspose.Cells textbox | C# Aspose.Cells multiline textbox alignment | how to change paragraph alignment inside a shape Aspose.Cells | Aspose.Cells TextBox paragraph formatting example
// Developer Intent: Apply distinct alignments to individual paragraphs within a TextBox shape.
// Use Cases: Design a report header where the title is centered and the subtitle is left‑aligned in the same textbox. | Create a data‑entry form with instructions left‑aligned and a section heading centered inside one textbox. | Generate an invoice note with a left‑aligned thank‑you line and a centered follow‑up line.
// AI Prompts: Show C# code that sets left alignment for the first paragraph and center alignment for the second paragraph of a TextBox using Aspose.Cells. | Provide an Aspose.Cells for .NET example that formats multiline textbox text with different paragraph alignments. | Explain how to safely check the paragraph count before applying alignment changes to a TextBox shape in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

// Creates a workbook, adds a TextBox shape, inserts two lines of text, and sets the first paragraph to left alignment and the second paragraph to center alignment using Aspose.Cells for .NET before saving the file.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add a textbox shape to the worksheet
        // Parameters: upper left row, upper left column, upper left X offset, upper left Y offset, width, height
        Shape textBox = sheet.Shapes.AddTextBox(0, 0, 100, 300, 200, 100);

        // Set multiline text (first line and second line)
        textBox.Text = "First line\nSecond line";

        // Get the collection of paragraphs inside the textbox
        TextParagraphCollection paragraphs = textBox.TextBody.TextParagraphs;

        // Apply left alignment to the first line (first paragraph)
        paragraphs[0].AlignmentType = TextAlignmentType.Left;

        // Apply center alignment to the second line (second paragraph) if it exists
        if (paragraphs.Count > 1)
        {
            paragraphs[1].AlignmentType = TextAlignmentType.Center;
        }

        // Save the workbook to a file
        workbook.Save("TextBoxAlignmentDemo.xlsx");
    }
}
