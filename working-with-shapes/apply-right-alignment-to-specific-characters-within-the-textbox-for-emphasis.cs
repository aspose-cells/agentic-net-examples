// Title: How to format individual characters inside an Excel TextBox using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set bold red style for a word inside a worksheet TextBox | format specific characters in an Excel TextBox with Aspose.Cells .NET API | apply italic blue formatting to a substring in a TextBox using Aspose.Cells | use Characters method to style parts of TextBox text in C# Aspose.Cells | change font properties of selected text in an Excel shape programmatically
// Tags: character-level formatting Aspose.Cells TextBox | partial text styling Excel shape C# | Aspose.Cells TextBox font customization | C# format substring in Excel TextBox | Aspose.Cells Characters method usage

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

// The example creates a workbook, adds a TextBox to the first worksheet, sets its text, and then uses the Characters method to apply bold red formatting to the word "Important" and italic blue formatting to the word "figures". Finally, the workbook is saved as an .xlsx file.
class Program
{
    static void Main()
    {
        const string outputPath = "RightAlignedCharactersInTextBox.xlsx";

        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a textbox to the worksheet (row, column, width, height in pixels)
            int textboxIndex = worksheet.TextBoxes.Add(2, 1, 160, 200);
            TextBox textBox = worksheet.TextBoxes[textboxIndex];

            if (textBox == null)
            {
                Console.WriteLine("Failed to create a TextBox.");
                return;
            }

            // Set the full text of the textbox
            textBox.Text = "Important: Please review the figures carefully.";

            // Note: Right alignment of the whole textbox is not required for the current task.
            // If needed, alignment can be adjusted via the TextBox's TextBody.TextAlignment properties.

            // Emphasise the word "Important" (bold and red)
            FontSetting importantPart = textBox.Characters(0, 9); // "Important"
            importantPart.Font.IsBold = true;
            importantPart.Font.Color = Color.Red;

            // Optionally format the word "figures"
            int figuresStart = textBox.Text.IndexOf("figures", StringComparison.Ordinal);
            if (figuresStart >= 0)
            {
                FontSetting figuresPart = textBox.Characters(figuresStart, "figures".Length);
                figuresPart.Font.IsItalic = true;
                figuresPart.Font.Color = Color.Blue;
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
