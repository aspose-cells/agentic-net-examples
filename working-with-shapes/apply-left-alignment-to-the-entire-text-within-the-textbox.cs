// Title: Aspose.Cells C# Example – Left Align Text in a TextBox Shape
// Description: Shows how to create a workbook, add a TextBox shape, insert multiline text, set each paragraph’s AlignmentType to TextAlignmentType.Left with Aspose.Cells for .NET, and save the workbook as an .xlsx file.
// Keywords: Aspose.Cells | C# textbox alignment | TextBox left alignment | TextParagraph AlignmentType | Aspose.Cells .NET example | shape text formatting | Excel left aligned text | Aspose.Cells GitHub | coding example | Aspose.Cells API
// Common Searches: left align text in Aspose.Cells textbox C# | Aspose.Cells set paragraph alignment .NET | how to align text inside a shape using Aspose.Cells | Aspose.Cells TextBox alignment example | C# code to left‑justify textbox paragraphs in Excel
// Developer Intent: Apply left alignment to every paragraph inside a TextBox shape when generating an Excel file with Aspose.Cells for .NET.
// Use Cases: Generating reports where notes in a textbox must be left‑aligned for readability. | Creating invoice templates with description fields inside a textbox that require consistent left alignment. | Automating formatting of user comments placed in a textbox to maintain a uniform appearance across exported workbooks.
// AI Prompts: Provide a C# snippet that left‑aligns all paragraphs in an Aspose.Cells TextBox shape and saves the workbook. | Explain how to modify the example to use center alignment instead of left alignment for textbox paragraphs. | Show how to retrieve the TextParagraphCollection from a Shape and set AlignmentType for each paragraph using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

namespace AsposeCellsExamples
{
    // Shows how to create a workbook, add a TextBox shape, insert multiline text, set each paragraph’s AlignmentType to TextAlignmentType.Left with Aspose.Cells for .NET, and save the workbook as an .xlsx file.
    public class TextBoxLeftAlignmentDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a text box shape to the worksheet
                // Parameters: upper left row, upper left column, top, left, width, height (in points)
                Shape textBox = worksheet.Shapes.AddTextBox(0, 0, 100, 300, 200, 100);

                // Set multiline text in the text box
                textBox.Text = "First Line\nSecond Line\nThird Line";

                // Retrieve all paragraphs in the text box
                TextParagraphCollection paragraphs = textBox.TextBody.TextParagraphs;

                // Apply left alignment to each paragraph
                foreach (TextParagraph paragraph in paragraphs)
                {
                    paragraph.AlignmentType = TextAlignmentType.Left;
                }

                // Save the workbook to a file
                string outputPath = "TextBoxLeftAlignmentDemo.xlsx";
                workbook.Save(outputPath);

                Console.WriteLine($"Text box with left-aligned text saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            TextBoxLeftAlignmentDemo.Run();
        }
    }
}
