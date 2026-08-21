// Title: Right‑Align TextBox Paragraph and Highlight Specific Characters with Aspose.Cells for .NET (C#)
// Description: This example creates a workbook, inserts a textbox shape, sets the paragraph alignment to right, and uses the FormatCharacters method with a StyleFlag to apply bold, italic and red color to a selected substring inside the textbox before saving the file.
// Keywords: Aspose.Cells right align textbox | FormatCharacters C# | StyleFlag text formatting | highlight substring in shape | textbox paragraph alignment .NET | apply bold italic color Aspose.Cells
// Common Searches: Aspose.Cells set textbox paragraph alignment to right | How to format part of text in a textbox shape using C# | Apply bold italic red style to specific characters Aspose.Cells | FormatCharacters method example for shapes | Right‑align text inside a shape with Aspose.Cells
// Developer Intent: Align a textbox paragraph to the right and style a chosen character range with bold, italic and color.
// Use Cases: Right‑aligned report footer where the word "Confidential" is emphasized in red, bold and italic. | Form label that aligns the instruction text to the right while highlighting the action word "Submit". | Dashboard widget with right‑aligned guidance and a highlighted alert term using custom font attributes.
// AI Prompts: Generate C# code with Aspose.Cells that adds a textbox, aligns its paragraph to the right, and formats the word "Important" in bold, italic, and blue. | Show how to format multiple non‑contiguous character ranges in a textbox shape with different styles using Aspose.Cells for .NET. | Demonstrate using StyleFlag and FormatCharacters to underline and color the substring "Review" green inside a right‑aligned textbox.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

namespace AsposeCellsExamples
{
    // This example creates a workbook, inserts a textbox shape, sets the paragraph alignment to right, and uses the FormatCharacters method with a StyleFlag to apply bold, italic and red color to a selected substring inside the textbox before saving the file.
    public class RightAlignSpecificCharactersDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a text box shape to the worksheet
                // Parameters: upper left row, upper left column, lower right row, lower right column, width, height
                Shape textBox = worksheet.Shapes.AddTextBox(0, 0, 100, 100, 300, 100);
                textBox.Text = "Important: Align this part to the right";

                // Set the paragraph alignment of the first paragraph to Right
                TextParagraph paragraph = textBox.TextBody.TextParagraphs[0];
                paragraph.AlignmentType = TextAlignmentType.Right;

                // Create a font for the characters we want to emphasize
                Aspose.Cells.Font emphasisFont = textBox.Font;
                emphasisFont.IsBold = true;          // make it bold
                emphasisFont.IsItalic = true;        // make it italic
                emphasisFont.Color = Color.Red;      // change color for visibility

                // Create a StyleFlag indicating which font properties to apply
                StyleFlag flag = new StyleFlag
                {
                    FontBold = true,
                    FontItalic = true,
                    FontColor = true
                };

                // Define the range of characters to emphasize (e.g., "Align this part")
                string text = textBox.Text;
                int startIndex = text.IndexOf("Align", StringComparison.Ordinal);
                if (startIndex >= 0)
                {
                    int length = "Align this part".Length;
                    // Apply formatting to the specific characters
                    textBox.FormatCharacters(startIndex, length, emphasisFont, flag);
                }

                // Save the workbook
                string outputPath = "RightAlignSpecificCharactersDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
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
            RightAlignSpecificCharactersDemo.Run();
        }
    }
}
