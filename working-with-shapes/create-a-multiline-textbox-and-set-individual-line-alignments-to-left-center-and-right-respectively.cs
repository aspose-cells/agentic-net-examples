// Title: C# – Create a multiline TextBox in Aspose.Cells and align each line left, center, or right
// Description: Demonstrates how to add a TextBox shape to an Excel worksheet with Aspose.Cells, insert three newline‑separated lines, retrieve the TextParagraphCollection, and set the AlignmentType of each paragraph to Left, Center, and Right before saving the workbook.
// Keywords: Aspose.Cells | C# | multiline TextBox | text alignment | paragraph alignment | TextBox shape | TextAlignmentType | Excel shape formatting | Aspose.Cells .NET example
// Common Searches: Aspose.Cells multiline textbox alignment C# | set different alignment for each line in a TextBox using Aspose.Cells | how to align paragraphs inside a TextBox shape in Excel with Aspose.Cells | C# Aspose.Cells left center right alignment per line
// Developer Intent: Add a TextBox shape with three lines of text and apply distinct horizontal alignments (left, center, right) to each line using Aspose.Cells for .NET.
// Use Cases: Create a report header where the title is centered, a subtitle left‑aligned, and a page number right‑aligned within a single textbox. | Design a product label that shows the address left‑aligned, the product name centered, and the price right‑aligned in an Excel worksheet. | Build an instructional sheet with steps aligned differently to improve visual hierarchy inside one textbox.
// AI Prompts: Show C# code that adds a multiline TextBox to an Aspose.Cells worksheet and sets left, center, and right alignment for each line. | Explain how to access the TextParagraphCollection of a TextBox shape and change the AlignmentType of individual paragraphs in Aspose.Cells for .NET. | Provide an example that customizes font style and horizontal alignment for each line of a multiline TextBox using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

// Demonstrates how to add a TextBox shape to an Excel worksheet with Aspose.Cells, insert three newline‑separated lines, retrieve the TextParagraphCollection, and set the AlignmentType of each paragraph to Left, Center, and Right before saving the workbook.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a textbox shape to the worksheet
        // Parameters: upper row, left column, height (pixels), width (pixels), upper offset, left offset
        Shape textBox = worksheet.Shapes.AddTextBox(0, 0, 100, 300, 200, 100);

        // Set multiline text (each line separated by newline)
        textBox.Text = "Left aligned line\nCenter aligned line\nRight aligned line";

        // Retrieve the collection of paragraphs (each line is a paragraph)
        TextParagraphCollection paragraphs = textBox.TextBody.TextParagraphs;

        // Apply individual horizontal alignments
        paragraphs[0].AlignmentType = TextAlignmentType.Left;    // First line: left
        paragraphs[1].AlignmentType = TextAlignmentType.Center;  // Second line: center
        paragraphs[2].AlignmentType = TextAlignmentType.Right;   // Third line: right

        // Save the workbook to a file
        workbook.Save("MultilineTextBoxAlignment.xlsx");
    }
}
