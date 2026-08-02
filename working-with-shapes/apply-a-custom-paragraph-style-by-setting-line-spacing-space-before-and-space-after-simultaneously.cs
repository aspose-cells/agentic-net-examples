// Title: Aspose.Cells .NET – Set line spacing, space before and after for a paragraph in a text box
// Description: This example creates a workbook, adds a text box with three paragraphs, selects the second paragraph, and configures its line spacing (points), space‑before, and space‑after values before saving the file.
// Keywords: Aspose.Cells paragraph line spacing | text box spacing Aspose.Cells .NET | set space before after paragraph | LineSpaceSizeType Points | TextParagraph formatting | C# Aspose.Cells shape text | Excel text box paragraph style | programmatic paragraph spacing | Aspose.Cells API paragraph properties | custom paragraph layout
// Common Searches: how to change line spacing of a paragraph in a text box using Aspose.Cells | Aspose.Cells set space before and after for a specific paragraph | modify TextParagraph spacing properties .NET | Aspose.Cells paragraph formatting inside shapes | C# set paragraph line spacing points Aspose.Cells
// Developer Intent: Apply line‑spacing, space‑before, and space‑after settings to a chosen paragraph inside a text‑box shape.
// Use Cases: Create a highlighted heading in a report by giving the second paragraph extra spacing. | Improve readability of multi‑paragraph notes in an Excel dashboard. | Standardize paragraph spacing across multiple shapes when generating worksheets programmatically.
// AI Prompts: Generate C# code that sets line spacing, space before, and space after for every paragraph in a text box with Aspose.Cells. | Show how to loop through a TextParagraphCollection and apply a 14‑point line spacing to each paragraph. | Explain how to retrieve a TextParagraph by index and modify its spacing attributes in Aspose.Cells .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

namespace AsposeCellsExamples
{
    // This example creates a workbook, adds a text box with three paragraphs, selects the second paragraph, and configures its line spacing (points), space‑before, and space‑after values before saving the file.
    public class CustomParagraphStyleDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a text box shape to the worksheet
                // Parameters: upper left row, upper left column, upper left offset X, upper left offset Y, width, height
                Shape textBox = worksheet.Shapes.AddTextBox(0, 0, 0, 0, 400, 200);

                // Set multi‑paragraph text (line break creates a new paragraph)
                textBox.Text = "First paragraph\nSecond paragraph\nThird paragraph";

                // Access the collection of paragraphs inside the text box
                TextParagraphCollection paragraphs = textBox.TextBody.TextParagraphs;

                // Choose the second paragraph (index 1) to apply custom spacing
                if (paragraphs.Count > 1)
                {
                    TextParagraph paragraph = paragraphs[1];

                    // Set line spacing type to points and define the line spacing value
                    paragraph.LineSpaceSizeType = LineSpaceSizeType.Points;
                    paragraph.LineSpace = 12; // 12 points line spacing

                    // Set space before and space after (also in points)
                    paragraph.SpaceBefore = 8;   // 8 points before the paragraph
                    paragraph.SpaceAfter = 10;   // 10 points after the paragraph
                }

                // Save the workbook to a file
                string outputPath = "CustomParagraphStyleDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            CustomParagraphStyleDemo.Run();
        }
    }
}
