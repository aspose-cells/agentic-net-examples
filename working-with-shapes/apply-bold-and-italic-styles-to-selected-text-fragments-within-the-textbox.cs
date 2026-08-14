// Title: Apply Bold and Italic Formatting to Text Inside a TextBox Shape with Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, insert a TextBox shape, set its text, and use the TextBody FontSettingCollection together with Style and StyleFlag objects to apply bold to the word “Bold” and italic to the word “Italic”. The workbook is then saved as an Excel file.
// Keywords: Aspose.Cells textbox bold | Aspose.Cells italic text | C# rich text in Excel shape | FontSettingCollection format substring | mixed font styles Aspose.Cells | apply style to part of textbox | Excel shape text formatting .NET
// Common Searches: how to make a word bold in an Aspose.Cells textbox | apply italic to part of textbox text Aspose.Cells C# | format substrings inside Excel shape using Aspose.Cells | Aspose.Cells FontSettingCollection example | C# set mixed font styles in Excel textbox
// Developer Intent: Use Aspose.Cells for .NET to apply distinct bold and italic styles to specific words within a TextBox shape in an Excel workbook.
// Use Cases: Emphasize key terms in a dashboard textbox. | Create a title with mixed bold/italic styling inside a shape. | Generate instructional notes where certain words need separate emphasis.
// AI Prompts: Write C# code that adds underline and a custom font color to the word "example" inside an Aspose.Cells textbox. | Show how to replace a specific word in all textboxes of a worksheet and apply bold formatting to each occurrence using Aspose.Cells for .NET. | Provide a method to apply both bold and italic styles simultaneously to a selected character range in a textbox shape.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

// Demonstrates how to create a workbook, insert a TextBox shape, set its text, and use the TextBody FontSettingCollection together with Style and StyleFlag objects to apply bold to the word “Bold” and italic to the word “Italic”. The workbook is then saved as an Excel file.
class ApplyBoldItalicInTextBox
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a textbox shape to the worksheet
        // Parameters: upper left row, upper left column, lower right row, lower right column, width, height
        Shape textBox = worksheet.Shapes.AddTextBox(0, 0, 0, 0, 300, 100);

        // Set the text of the textbox
        textBox.Text = "Bold and Italic text example";

        // Get the FontSettingCollection that manages rich text inside the textbox
        FontSettingCollection textBody = textBox.TextBody;

        // ---------- Apply Bold to the word "Bold" ----------
        int boldStart = 0;                     // start index of "Bold"
        int boldLength = "Bold".Length;        // length of "Bold"

        // Create a style with Bold enabled
        Style boldStyle = workbook.CreateStyle();
        boldStyle.Font.IsBold = true;

        // Define which font properties to apply (only Bold)
        StyleFlag boldFlag = new StyleFlag();
        boldFlag.FontBold = true;

        // Apply the bold formatting to the specified range
        textBody.Format(boldStart, boldLength, boldStyle.Font, boldFlag);

        // ---------- Apply Italic to the word "Italic" ----------
        string fullText = textBox.Text;
        int italicStart = fullText.IndexOf("Italic");   // locate "Italic"
        int italicLength = "Italic".Length;

        // Create a style with Italic enabled
        Style italicStyle = workbook.CreateStyle();
        italicStyle.Font.IsItalic = true;

        // Define which font properties to apply (only Italic)
        StyleFlag italicFlag = new StyleFlag();
        italicFlag.FontItalic = true;

        // Apply the italic formatting to the specified range
        textBody.Format(italicStart, italicLength, italicStyle.Font, italicFlag);

        // Save the workbook to a file
        workbook.Save("BoldItalicTextBox.xlsx");
    }
}
