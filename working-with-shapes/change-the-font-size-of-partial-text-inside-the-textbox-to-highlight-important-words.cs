using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

namespace AsposeCellsPartialTextFormatting
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a textbox shape to the worksheet
            // Parameters: upper left row, upper left column, upper left offset X, upper left offset Y, width, height
            Shape textBox = worksheet.Shapes.AddTextBox(1, 1, 0, 0, 300, 100);
            textBox.Text = "Important: Review the quarterly report";

            // Define the portion of text to highlight (e.g., the word "Important")
            int startIndex = 0;               // start at the first character
            int length = "Important".Length;  // length of the word to format

            // Create a Font object based on the textbox's current font
            Font highlightFont = textBox.Font;
            highlightFont.Size = 20;          // set desired font size for the highlighted part

            // Create a StyleFlag indicating which font attributes to apply
            StyleFlag flag = new StyleFlag();
            flag.FontSize = true;             // we only change the font size in this example

            // Apply the formatting to the specified characters
            textBox.FormatCharacters(startIndex, length, highlightFont, flag);

            // Save the workbook to a file
            workbook.Save("PartialTextHighlight.xlsx");
        }
    }
}