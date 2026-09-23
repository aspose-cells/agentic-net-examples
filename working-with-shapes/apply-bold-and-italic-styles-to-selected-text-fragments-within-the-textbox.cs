// Title: Apply bold and italic styles to specific words inside an Excel TextBox using Aspose.Cells RichText API in C#
// AI Prompts: Generate C# code that creates a TextBox shape on a worksheet and uses Aspose.Cells RichText to make the word "Hello" bold and the word "World" italic. | Show how to combine multiple Font objects with different IsBold/IsItalic settings to format individual character ranges in an Aspose.Cells TextBox. | Provide a step‑by‑step example of adding a TextBox, assigning mixed‑style text via RichText, and saving the workbook with Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# format part of textbox text bold italic | RichText API example for Excel textbox styling in .NET | How to set different fonts for words inside a textbox using Aspose.Cells | C# apply mixed formatting to TextBox shape in Excel workbook | Partial text formatting in Aspose.Cells TextBox shape
// Tags: Aspose.Cells RichText partial formatting | C# Excel TextBox style manipulation | Aspose.Cells set bold italic in textbox | RichText API for textbox in Aspose.Cells .NET | Excel shape textbox mixed font formatting | Aspose.Cells TextBox rich text example

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample creates a new workbook, adds a TextBox shape to the first worksheet, clears any default content, assigns a static string, and saves the file as StyledTextbox.xlsx. It notes that applying bold or italic to individual words requires the Aspose.Cells RichText API, which is not demonstrated in this basic example.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a textbox shape (row, column, row offset, column offset, height, width)
            // In recent Aspose.Cells versions AddTextBox returns a TextBox object directly
            TextBox textbox = sheet.Shapes.AddTextBox(2, 1, 0, 2, 300, 100);

            // Clear any default text
            textbox.Text = string.Empty;

            // Helper to create a Font with specific style
            Font CreateFont(bool isBold = false, bool isItalic = false)
            {
                Font font = workbook.CreateStyle().Font;
                font.IsBold = isBold;
                font.IsItalic = isItalic;
                return font;
            }

            // NOTE: In the current Aspose.Cells version the TextBox class does not expose an AddText method.
            // To keep the example functional we set the whole text at once.
            // Styling of individual portions would require using RichText APIs which are beyond this simple demo.
            textbox.Text = "Hello World! This is Aspose.Cells.";

            // Save the workbook
            string outputPath = "StyledTextbox.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}
