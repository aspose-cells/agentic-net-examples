// Title: How to left‑align text inside a TextBox shape using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells set left horizontal alignment for a TextBox shape | apply left alignment to textbox content in an Excel workbook with Aspose.Cells | Aspose.Cells TextHorizontalAlignment property example in C# | auto‑size a TextBox after aligning its text left using Aspose.Cells .NET | left align all text in a worksheet TextBox via Aspose.Cells API
// Tags: Aspose.Cells TextBox left alignment C# | TextHorizontalAlignment property Aspose.Cells | auto size TextBox content Aspose.Cells | Excel textbox shape alignment .NET | Aspose.Cells shape text formatting

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

namespace ApplyLeftAlignmentToTextBox
{
    // The sample creates a new workbook, adds a TextBox shape to the first worksheet, sets sample text, applies left horizontal alignment using the TextHorizontalAlignment property, enables auto‑size for the text body, and saves the file as an XLSX workbook.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a textbox to the worksheet (row, column, width, height)
                int textboxIndex = worksheet.TextBoxes.Add(2, 1, 160, 200);
                TextBox textBox = worksheet.TextBoxes[textboxIndex];

                // Set sample text
                textBox.Text = "This is a sample text that will be left‑aligned inside the textbox.";

                // Apply left alignment to the entire text within the textbox
                // TextBox inherits from Shape, which provides TextHorizontalAlignment property
                textBox.TextHorizontalAlignment = TextAlignmentType.Left;

                // Optional: enable auto‑size so the textbox fits its content
                textBox.TextBody.TextAlignment.AutoSize = true;

                // Save the workbook
                string outputPath = "LeftAlignedTextBox.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);

                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while applying left alignment to the textbox:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
