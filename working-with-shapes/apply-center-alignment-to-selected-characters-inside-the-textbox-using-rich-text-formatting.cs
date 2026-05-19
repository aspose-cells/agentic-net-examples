using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

namespace AsposeCellsExamples
{
    public class CenterAlignSelectedCharactersDemo
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
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a text box shape to the worksheet
            // Parameters: upper left row, upper left column, upper left offset, upper left offset,
            // width (in points), height (in points)
            Shape textBox = worksheet.Shapes.AddTextBox(0, 0, 0, 0, 300, 100);

            // Set the text of the textbox
            textBox.Text = "Aspose.Cells Rich Text Formatting Example";

            // Get the FontSettingCollection which represents the text body of the shape
            FontSettingCollection textBody = textBox.TextBody;

            // Define the range of characters we want to treat as a separate "rich text" segment
            // For example, characters 7 to 12 ("Cells")
            int startIndex = 7;
            int length = 5;

            // Create a temporary style to apply (here we just change the font color to illustrate)
            Style style = workbook.CreateStyle();
            style.Font.Color = Color.Blue;
            StyleFlag flag = new StyleFlag();
            flag.FontColor = true;

            // Apply the style to the selected characters
            textBody.Format(startIndex, length, style.Font, flag);

            // Align the paragraph that contains the formatted characters to center
            TextParagraph paragraph = textBody.TextParagraphs[0];
            paragraph.AlignmentType = TextAlignmentType.Center;

            // Save the workbook
            string outputPath = "CenterAlignSelectedCharactersDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
    }
}