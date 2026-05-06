using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

class ReplaceWithFormattingDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a text box shape to the worksheet
        Shape textBox = worksheet.Shapes.AddTextBox(0, 0, 100, 100, 300, 100);

        // Get the FontSettingCollection (text body) of the shape
        FontSettingCollection fontSettings = textBox.TextBody;

        // Set initial text in the text box
        fontSettings.Text = "Hello Aspose Cells!";

        // Apply formatting to the word "Aspose"
        Style style = workbook.CreateStyle();
        style.Font.Name = "Arial";
        style.Font.Size = 14;
        style.Font.IsBold = true;
        style.Font.Color = Color.Blue;

        // Determine the start index of the word to format
        int startIndex = fontSettings.Text.IndexOf("Aspose");
        int length = "Aspose".Length;

        // Create a StyleFlag indicating which font properties to apply
        StyleFlag flag = new StyleFlag
        {
            FontName = true,
            FontSize = true,
            FontBold = true,
            FontColor = true
        };

        // Apply the formatting to the specified characters
        fontSettings.Format(startIndex, length, style.Font, flag);

        // Replace the word "Aspose" with "World" while preserving existing formatting
        fontSettings.Replace("Aspose", "World");

        // Save the workbook to a file
        workbook.Save("ReplaceWithFormattingDemo.xlsx");
    }
}