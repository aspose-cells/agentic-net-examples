using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

namespace AsposeCellsExamples
{
    public class SetSpaceBeforeDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a text box shape to the worksheet
                // Parameters: upper left row, upper left column, upper left offset X, upper left offset Y, width, height
                Shape textBox = worksheet.Shapes.AddTextBox(0, 0, 0, 0, 400, 200);

                // Set text with multiple paragraphs (separated by newline)
                textBox.Text = "First paragraph\nSecond paragraph\nThird paragraph";

                // Access the collection of paragraphs in the text box
                TextParagraphCollection paragraphs = textBox.TextBody.TextParagraphs;

                // Apply a spacing of 3 points before each paragraph
                foreach (TextParagraph paragraph in paragraphs)
                {
                    paragraph.SpaceBefore = 3.0; // 3 points
                }

                // Define output file path
                string outputPath = "SetSpaceBeforeDemo.xlsx";

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            SetSpaceBeforeDemo.Run();
        }
    }
}