// Title: Aspose.Cells .NET – Add a TextBox shape and set its text with TextBody.Text
// Description: Demonstrates how to create a Workbook, insert a TextBox shape on the first worksheet, assign text through the TextBody.Text property (FontSettingCollection), and save the file as an XLSX document using Aspose.Cells for .NET.
// Keywords: Aspose.Cells add textbox | TextBox TextBody.Text | FontSettingCollection text property | C# Aspose.Cells shape text | set textbox caption Aspose.Cells | save workbook with textbox
// Common Searches: Aspose.Cells add textbox shape C# | How to set TextBox text using TextBody.Text in Aspose.Cells | C# example for TextBox TextBody property Aspose.Cells | Saving Excel file after inserting textbox Aspose.Cells | FontSettingCollection Text property Aspose.Cells
// Developer Intent: Insert a TextBox shape into a worksheet and define its displayed text via the TextBody.Text property.
// Use Cases: Add labeled annotations to generated Excel reports. | Create dynamic header or title boxes that display variable strings. | Populate multiple textboxes with data‑driven content during export.
// AI Prompts: Write C# code that adds a TextBox shape to an Aspose.Cells worksheet, sets its caption using TextBody.Text, and saves the workbook. | Show how to modify the TextBody.Text of an existing TextBox in an Aspose.Cells workbook. | Explain how to apply font size, color, and style to a TextBox after setting TextBody.Text with FontSettingCollection.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

// Demonstrates how to create a Workbook, insert a TextBox shape on the first worksheet, assign text through the TextBody.Text property (FontSettingCollection), and save the file as an XLSX document using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a textbox shape to the worksheet
        // Parameters: topRow, top (pixel), leftColumn, left (pixel), height (pixel), width (pixel)
        TextBox textBox = worksheet.Shapes.AddTextBox(1, 0, 1, 0, 100, 200);

        // Set the text using the TextBody.Text property (FontSettingCollection.Text)
        FontSettingCollection fontSettings = textBox.TextBody;
        fontSettings.Text = "Hello, Aspose.Cells!";

        // Save the workbook
        workbook.Save("TextBoxWithTextBody.xlsx");
    }
}
