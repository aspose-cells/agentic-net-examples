using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

namespace AsposeCellsExamples
{
    public class BalancedTextAppearanceDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a text box shape to the worksheet
                TextBox textBox = worksheet.Shapes.AddTextBox(0, 0, 0, 0, 300, 150);
                textBox.Text = "First line of text\nSecond line of text\nThird line of text";

                // Adjust character spacing for the whole text run
                TextOptions textOptions = textBox.TextOptions;
                textOptions.Spacing = 1.5; // Positive value increases spacing between characters

                // Access the first paragraph to set line spacing and paragraph spacing
                TextParagraph paragraph = textBox.TextBody.TextParagraphs[0];
                paragraph.LineSpaceSizeType = LineSpaceSizeType.Points;
                paragraph.LineSpace = 8; // Increase vertical space between lines (points)

                // Optional: add space before and after the paragraph for better visual balance
                paragraph.SpaceBefore = 4; // Space before the paragraph (points)
                paragraph.SpaceAfter = 4;  // Space after the paragraph (points)

                // Save the workbook to a file
                string outputPath = "BalancedTextAppearance.xlsx";
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
            BalancedTextAppearanceDemo.Run();
        }
    }
}