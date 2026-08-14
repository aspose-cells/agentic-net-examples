// Title: Highlight partial text in an Aspose.Cells textbox – change font size, bold, and color (C#)
// Description: Shows how to create a workbook with Aspose.Cells, add a textbox shape, and use Font, StyleFlag, and FormatCharacters to enlarge, bold, and color a specific word inside the textbox before saving the file.
// Keywords: Aspose.Cells C# textbox format characters | Aspose.Cells change font size partial text | Aspose.Cells StyleFlag usage | FormatCharacters method example | highlight word in Excel shape | C# Excel textbox styling | partial text formatting Aspose.Cells | Excel shape text formatting .NET | Aspose.Cells Font properties | Aspose.Cells text box partial formatting
// Common Searches: Aspose.Cells change font size of a word in textbox | How to bold part of text in Excel shape using C# | FormatCharacters example Aspose.Cells | Apply color to specific text in Aspose.Cells textbox | Partial text formatting in Aspose.Cells .NET
// Developer Intent: Apply custom font size, weight, and color to a selected substring within a textbox shape in an Excel workbook using Aspose.Cells for .NET.
// Use Cases: Emphasize key terms in a report header by enlarging and coloring them inside a textbox. | Create warning or instruction boxes where critical words appear in bold red while the rest of the text stays normal. | Generate dynamic dashboards that highlight changing metrics within shape annotations without altering the whole textbox.
// AI Prompts: Show how to format multiple distinct words inside an Aspose.Cells textbox with different font sizes and colors using C#. | Provide a C# example that underlines and italicizes a phrase in a shape while leaving the remaining text unchanged. | Explain step‑by‑step how StyleFlag and FormatCharacters work together to apply bold, color, and size to a substring in an Aspose.Cells textbox.

using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Shows how to create a workbook with Aspose.Cells, add a textbox shape, and use Font, StyleFlag, and FormatCharacters to enlarge, bold, and color a specific word inside the textbox before saving the file.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a textbox shape to the worksheet
        // Parameters: upper left row, upper left column, top, left, width, height
        Shape textBox = worksheet.Shapes.AddTextBox(0, 0, 100, 100, 200, 50);
        textBox.Text = "Important: Review the quarterly report";

        // Prepare a font with the desired highlight properties
        Font highlightFont = textBox.Font;
        highlightFont.Size = 20;          // larger font size
        highlightFont.IsBold = true;      // bold
        highlightFont.Color = Color.Red; // red color

        // Specify which font attributes should be applied
        StyleFlag flag = new StyleFlag();
        flag.FontSize = true;
        flag.FontBold = true;
        flag.FontColor = true;

        // Determine the start index and length of the word to highlight
        string fullText = textBox.Text;
        string wordToHighlight = "Important";
        int startIndex = fullText.IndexOf(wordToHighlight);
        int length = wordToHighlight.Length;

        // Apply the formatting to the selected characters
        textBox.FormatCharacters(startIndex, length, highlightFont, flag);

        // Save the workbook
        workbook.Save("HighlightedTextbox.xlsx");
    }
}
