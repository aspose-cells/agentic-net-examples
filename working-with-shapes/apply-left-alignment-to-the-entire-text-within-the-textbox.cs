// Title: Left‑align all paragraphs in a textbox shape using Aspose.Cells for .NET (C#)
// Description: Creates a workbook, adds a textbox shape, inserts multiline text, iterates through the TextParagraphCollection, sets each paragraph's AlignmentType to TextAlignmentType.Left, and saves the file as an .xlsx document.
// Keywords: Aspose.Cells C# textbox alignment | set left alignment Aspose.Cells | TextParagraph AlignmentType Left | Excel shape text formatting .NET | Aspose.Cells TextBox example
// Common Searches: Aspose.Cells left align textbox text C# | How to set paragraph alignment in an Excel shape using Aspose.Cells | C# code to left‑justify text in a textbox shape | Change textbox paragraph alignment Aspose.Cells .NET
// Developer Intent: The developer needs to left‑align every paragraph inside a textbox shape in an Excel workbook using Aspose.Cells for .NET.
// Use Cases: Design reports where textbox captions must be left‑justified for readability. | Build invoice templates with left‑aligned notes inside multiline textboxes. | Automate worksheet templates that require consistent left alignment of textbox content.
// AI Prompts: Show how to left‑align all paragraphs in a textbox shape with Aspose.Cells for .NET (C#). | Provide a C# snippet to change a textbox's text alignment to right or center using Aspose.Cells. | Explain how to access and modify TextParagraph objects in a shape's TextBody with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

namespace AsposeCellsExamples
{
    // Creates a workbook, adds a textbox shape, inserts multiline text, iterates through the TextParagraphCollection, sets each paragraph's AlignmentType to TextAlignmentType.Left, and saves the file as an .xlsx document.
    public class TextBoxLeftAlignmentDemo
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
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a text box shape to the worksheet
            // Parameters: upper left row, upper left column, upper left offset (pixels),
            // lower right row, lower right column, lower right offset (pixels)
            Shape textBox = worksheet.Shapes.AddTextBox(0, 0, 100, 300, 200, 100);

            // Set multiline text inside the text box
            textBox.Text = "First Line\nSecond Line\nThird Line";

            // Retrieve all paragraphs of the text box
            TextParagraphCollection paragraphs = textBox.TextBody.TextParagraphs;

            // Apply left alignment to each paragraph
            foreach (TextParagraph paragraph in paragraphs)
            {
                paragraph.AlignmentType = TextAlignmentType.Left;
            }

            // Save the workbook
            string outputPath = "TextBoxLeftAlignmentDemo.xlsx";
            workbook.Save(outputPath);

            Console.WriteLine($"Text box with left-aligned text saved successfully to '{outputPath}'.");
        }
    }
}
