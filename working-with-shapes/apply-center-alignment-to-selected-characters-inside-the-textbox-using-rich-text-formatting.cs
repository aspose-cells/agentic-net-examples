// Title: Center Align a Specific Paragraph in a TextBox Using Aspose.Cells for .NET (C#)
// Description: This C# example shows how to add a TextBox shape to a worksheet, insert multiline text, and set left, center, and right alignments on individual paragraphs via the TextParagraphCollection, then save the workbook as an XLSX file.
// Keywords: Aspose.Cells C# textbox alignment | center paragraph Aspose.Cells | TextParagraphCollection alignment | shape text formatting Aspose.Cells | rich text alignment Excel C# | Aspose.Cells paragraph formatting
// Common Searches: Aspose.Cells align paragraph in textbox | C# set center alignment for specific line in Excel shape | How to format text inside a textbox with Aspose.Cells | Center text in a textbox paragraph Aspose.Cells .NET | Apply different alignments to textbox lines using Aspose.Cells
// Developer Intent: Apply center alignment to selected characters (a specific paragraph) inside a textbox shape in an Excel workbook using Aspose.Cells for .NET.
// Use Cases: Create a report where the title line in a textbox is centered while other lines remain left‑aligned or right‑aligned. | Generate a spreadsheet with annotated text boxes that require distinct alignments for each paragraph. | Build a dashboard worksheet where only certain textbox paragraphs need to be centered for visual emphasis.
// AI Prompts: Show me C# code to center align a specific paragraph inside a textbox using Aspose.Cells. | Provide an Aspose.Cells example that applies left, center, and right alignments to different lines of a textbox. | Explain how to use TextParagraphCollection to format individual paragraphs in a shape with Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

namespace AsposeCellsExamples
{
    // This C# example shows how to add a TextBox shape to a worksheet, insert multiline text, and set left, center, and right alignments on individual paragraphs via the TextParagraphCollection, then save the workbook as an XLSX file.
    public class CenterAlignSelectedCharactersInTextBox
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a text box shape to the worksheet
            // Parameters: upper left row, upper left column, upper left offset X, upper left offset Y,
            // width, height (all in points)
            Shape textBox = worksheet.Shapes.AddTextBox(0, 0, 0, 0, 300, 150);

            // Set the text of the text box using line breaks for separate paragraphs
            textBox.Text = "First line of text\nCentered characters here\nThird line of text";

            // Access the collection of paragraphs inside the text box
            TextParagraphCollection paragraphs = textBox.TextBody.TextParagraphs;

            // Apply center alignment to the second paragraph (index 1)
            paragraphs[1].AlignmentType = TextAlignmentType.Center;

            // Optionally set other paragraphs to different alignments
            paragraphs[0].AlignmentType = TextAlignmentType.Left;
            paragraphs[2].AlignmentType = TextAlignmentType.Right;

            // Save the workbook
            string outputPath = "CenterAlignedCharactersInTextBox.xlsx";
            workbook.Save(outputPath);

            Console.WriteLine($"Workbook saved as '{outputPath}'.");
        }
    }
}
