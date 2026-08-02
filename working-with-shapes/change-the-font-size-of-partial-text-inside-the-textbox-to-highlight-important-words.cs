// Title: Change Font Size of Partial Text in a TextBox Shape with Aspose.Cells for .NET
// Description: Shows how to create a workbook, insert a TextBox shape, and apply the FormatCharacters method with a Font object and StyleFlag to enlarge only the word “Important” while leaving the remaining text unchanged, then save the result as PartialTextHighlight.xlsx.
// Keywords: Aspose.Cells | C# partial text formatting | textbox shape font size | FormatCharacters method | Excel shape text styling | highlight word in textbox | Aspose.Cells .NET | Excel partial font change | text box character formatting | shape text formatting
// Common Searches: increase font size of a word inside a textbox using Aspose.Cells | Aspose.Cells FormatCharacters example C# | how to format partial text in Excel shape | change font size of substring in Aspose.Cells textbox | partial text styling with Aspose.Cells .NET
// Developer Intent: Enlarge a specific substring inside a TextBox shape in an Excel workbook.
// Use Cases: Emphasize key terms in report annotations. | Create visual hierarchy in dashboard shapes. | Highlight dynamic values in chart callouts based on data.
// AI Prompts: Generate C# code that changes the font color of a selected range of characters in a TextBox shape using Aspose.Cells. | Provide an example that applies bold and italic styles to multiple text segments within an Excel shape. | Explain how to retrieve the Font object for a substring and modify its attributes with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsPartialTextFormatting
{
    // Shows how to create a workbook, insert a TextBox shape, and apply the FormatCharacters method with a Font object and StyleFlag to enlarge only the word “Important” while leaving the remaining text unchanged, then save the result as PartialTextHighlight.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a textbox shape to the worksheet
            // Parameters: upper left row, upper left column, upper left offset, upper left offset,
            // width, height (all in points)
            Shape textBox = worksheet.Shapes.AddTextBox(1, 0, 0, 100, 200, 50);
            textBox.Text = "Important: Review the quarterly results";

            // Define the portion of text to highlight (e.g., the word "Important")
            int startIndex = 0;               // start at the first character
            int length = "Important".Length;  // length of the word to format

            // Prepare a Font object with the desired size
            Aspose.Cells.Font highlightFont = textBox.Font;
            highlightFont.Size = 18; // larger size to highlight

            // Specify which font attributes should be applied
            StyleFlag flag = new StyleFlag();
            flag.FontSize = true; // only change the font size

            // Apply the formatting to the selected characters
            textBox.FormatCharacters(startIndex, length, highlightFont, flag);

            // Save the workbook
            workbook.Save("PartialTextHighlight.xlsx");
        }
    }
}
