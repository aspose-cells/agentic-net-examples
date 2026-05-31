using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    public class TextBoxBoldItalicDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a textbox shape to the worksheet
                // Parameters: upper left row, upper left column, top, left, width, height
                Shape textBox = worksheet.Shapes.AddTextBox(0, 0, 0, 0, 250, 80);
                textBox.Text = "Bold and Italic text example";

                // ---------- Apply Bold to the word "Bold" ----------
                // Create a font based on the textbox's current font
                Aspose.Cells.Font boldFont = textBox.Font;
                boldFont.IsBold = true; // set bold property

                // Define which font attributes to apply (only bold in this case)
                StyleFlag boldFlag = new StyleFlag { FontBold = true };

                // Apply formatting to characters starting at index 0 with length 4 ("Bold")
                textBox.FormatCharacters(0, 4, boldFont, boldFlag);

                // ---------- Apply Italic to the word "Italic" ----------
                // Create a separate font for italic formatting
                Aspose.Cells.Font italicFont = textBox.Font;
                italicFont.IsItalic = true; // set italic property

                // Define style flag for italic
                StyleFlag italicFlag = new StyleFlag { FontItalic = true };

                // "Italic" starts at index 9 and has length 6
                textBox.FormatCharacters(9, 6, italicFont, italicFlag);

                // Determine output file path
                string outputPath = "TextBoxBoldItalicDemo.xlsx";

                // Save the workbook to a file
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
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
            TextBoxBoldItalicDemo.Run();
        }
    }
}