// Title: C# – Apply Bold and Italic Formatting to Words in an Aspose.Cells TextBox Shape
// Description: Creates a workbook, inserts a TextBox shape, sets its text, and uses FontSettingCollection with Style and StyleFlag to make the word “Bold” bold and the word “Italic” italic, then saves the file as BoldItalicTextBox.xlsx.
// Keywords: Aspose.Cells C# textbox formatting | apply bold to shape text Aspose.Cells | apply italic to shape text Aspose.Cells | FontSettingCollection partial formatting | StyleFlag text styling Excel | C# Excel shape text styling | Aspose.Cells partial character formatting | Excel automation text styling
// Common Searches: how to bold part of a textbox in Aspose.Cells C# | italic text inside a shape using Aspose.Cells | C# format specific characters in Aspose.Cells TextBox | Aspose.Cells StyleFlag example for partial text | apply mixed font styles to Excel shape text
// Developer Intent: Apply bold and italic styles to individual words inside a TextBox shape in an Excel workbook using Aspose.Cells for .NET.
// Use Cases: Generate a sales report where key terms inside a textbox need distinct emphasis. | Create instructional worksheets that highlight action verbs with bold or italic styling programmatically. | Automate marketing templates that format product names differently within shape captions.
// AI Prompts: Show how to underline and change the color of a specific word in an Aspose.Cells TextBox using C#. | Provide a C# snippet that applies three different font styles to non‑contiguous words in a shape's TextBody. | Explain how to retrieve a TextBox's FontSettingCollection, modify alignment, line spacing, and then apply the changes.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

// Creates a workbook, inserts a TextBox shape, sets its text, and uses FontSettingCollection with Style and StyleFlag to make the word “Bold” bold and the word “Italic” italic, then saves the file as BoldItalicTextBox.xlsx.
class ApplyBoldItalicInTextBox
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a textbox shape to the worksheet
        // Parameters: upper left row, upper left column, top, left, width, height
        Shape textBox = worksheet.Shapes.AddTextBox(1, 1, 0, 0, 200, 50);

        // Set the initial text of the textbox
        textBox.Text = "Bold and Italic text";

        // Get the FontSettingCollection that represents the text body of the shape
        FontSettingCollection textBody = textBox.TextBody;

        // -------------------------------------------------
        // Apply Bold to the word "Bold"
        // -------------------------------------------------
        // Create a style to hold the font settings
        Style boldStyle = workbook.CreateStyle();
        boldStyle.Font.IsBold = true;          // set bold
        // Define which font properties should be applied
        StyleFlag boldFlag = new StyleFlag();
        boldFlag.FontBold = true;

        // Apply formatting: start index 0, length 4 ("Bold")
        textBody.Format(0, 4, boldStyle.Font, boldFlag);

        // -------------------------------------------------
        // Apply Italic to the word "Italic"
        // -------------------------------------------------
        Style italicStyle = workbook.CreateStyle();
        italicStyle.Font.IsItalic = true;      // set italic
        StyleFlag italicFlag = new StyleFlag();
        italicFlag.FontItalic = true;

        // "Italic" starts after "Bold and " (9 characters)
        int italicStart = "Bold and ".Length; // 9
        int italicLength = "Italic".Length;   // 6
        textBody.Format(italicStart, italicLength, italicStyle.Font, italicFlag);

        // Save the workbook to a file
        workbook.Save("BoldItalicTextBox.xlsx");
    }
}
