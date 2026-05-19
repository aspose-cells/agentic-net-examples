using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

namespace AsposeCellsParagraphStyleDemo
{
    public class ApplyCustomParagraphStyle
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

                // Choose the second paragraph (index 1) to apply the custom style
                TextParagraph paragraph = paragraphs[1];

                // Set line spacing unit to points and define the spacing value
                paragraph.LineSpaceSizeType = LineSpaceSizeType.Points;
                paragraph.LineSpace = 12; // 12 points line spacing

                // Set space before and after the paragraph (also in points)
                paragraph.SpaceBeforeSizeType = LineSpaceSizeType.Points;
                paragraph.SpaceBefore = 8;   // 8 points before the paragraph

                paragraph.SpaceAfterSizeType = LineSpaceSizeType.Points;
                paragraph.SpaceAfter = 10;   // 10 points after the paragraph

                // Save the workbook
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

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            ApplyCustomParagraphStyle.Run();
        }
    }
}