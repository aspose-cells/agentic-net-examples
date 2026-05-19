using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

namespace AsposeCellsExamples
{
    public class SetSpaceAfterDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a text box shape to the worksheet
                Shape textBox = worksheet.Shapes.AddTextBox(0, 0, 0, 0, 400, 200);
                // Set text with multiple paragraphs (separated by newline)
                textBox.Text = "First paragraph\nSecond paragraph\nThird paragraph";

                // Access the collection of paragraphs in the text box
                TextParagraphCollection paragraphs = textBox.TextBody.TextParagraphs;

                // Apply spacing after each paragraph: 4 points
                foreach (TextParagraph paragraph in paragraphs)
                {
                    paragraph.SpaceAfterSizeType = LineSpaceSizeType.Points;
                    paragraph.SpaceAfter = 4.0;
                }

                // Save the workbook to a file
                string outputPath = "SetSpaceAfterDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            SetSpaceAfterDemo.Run();
        }
    }
}