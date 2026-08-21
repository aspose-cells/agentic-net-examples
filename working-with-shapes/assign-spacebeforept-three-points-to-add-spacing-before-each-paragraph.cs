// Title: Add 3‑point SpaceBefore to every paragraph in a textbox shape with Aspose.Cells for .NET (C#)
// Description: Creates a new workbook, inserts a textbox shape on the first worksheet, fills it with three lines of text, iterates the TextParagraphCollection, and sets the SpaceBefore property to 3 points for each paragraph before saving the file as an Excel workbook.
// Keywords: Aspose.Cells | C# | .NET | Excel textbox shape | TextParagraph spacing | SpaceBefore property | paragraph leading | AddTextBox | TextParagraphCollection
// Common Searches: Aspose.Cells set SpaceBefore for textbox paragraphs | C# add space before paragraphs in Excel shape | How to adjust paragraph spacing in Aspose.Cells textbox | SpaceBefore points Aspose.Cells .NET
// Developer Intent: Apply a 3‑point space before each paragraph inside a textbox shape in an Excel workbook using Aspose.Cells for .NET.
// Use Cases: Design reports where textbox paragraphs need consistent leading for readability. | Build Excel templates with styled text boxes that follow corporate spacing standards. | Automate generation of documentation sheets where paragraph spacing must be uniform.
// AI Prompts: Show how to set different SpaceBefore values for individual paragraphs in an Aspose.Cells textbox. | Provide code to modify both SpaceBefore and SpaceAfter for paragraphs in a shape using Aspose.Cells for .NET. | Explain how to combine paragraph spacing with font style and alignment in Aspose.Cells text boxes.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

// Creates a new workbook, inserts a textbox shape on the first worksheet, fills it with three lines of text, iterates the TextParagraphCollection, and sets the SpaceBefore property to 3 points for each paragraph before saving the file as an Excel workbook.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a text box shape to the worksheet
        // Parameters: upper left row, upper left column, upper left offset X, upper left offset Y, width, height
        Shape textBox = worksheet.Shapes.AddTextBox(0, 0, 0, 0, 400, 200);

        // Set multiline text (each line is a separate paragraph)
        textBox.Text = "First paragraph\nSecond paragraph\nThird paragraph";

        // Retrieve the collection of paragraphs inside the text box
        TextParagraphCollection paragraphs = textBox.TextBody.TextParagraphs;

        // Assign 3 points of space before each paragraph
        foreach (TextParagraph paragraph in paragraphs)
        {
            paragraph.SpaceBefore = 3; // SpaceBefore is measured in points
        }

        // Save the workbook to a file
        workbook.Save("SpaceBeforeDemo.xlsx");
    }
}
