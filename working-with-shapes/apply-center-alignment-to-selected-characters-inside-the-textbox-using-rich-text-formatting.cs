// Title: C# – Center‑Align Selected Characters in an Aspose.Cells TextBox via Rich‑Text Formatting
// Description: Creates a workbook, adds a TextBox shape, applies a blue Arial style to the first five characters, centers the paragraph that contains those characters, and saves the file as an Excel workbook using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | textbox shape | center alignment | rich text formatting | FontSettingCollection | TextParagraph | partial text styling | Excel shape formatting | paragraph alignment
// Common Searches: Aspose.Cells center align part of textbox text | C# format selected characters in Aspose.Cells shape | rich‑text alignment in Aspose.Cells textbox | how to apply style to a range of characters in a textbox using Aspose.Cells | center paragraph inside a textbox Aspose.Cells .NET
// Developer Intent: Apply center alignment to a specific range of characters inside a textbox shape using Aspose.Cells rich‑text APIs.
// Use Cases: Generate a report where the header word inside a textbox is styled and centered while the remainder stays unchanged. | Build dynamic Excel dashboards that require different font styles for portions of a textbox and centered alignment for those portions. | Automate workbook creation that formats and aligns selected characters in shapes for branding or visual emphasis.
// AI Prompts: Provide C# code that uses FontSettingCollection to center‑align characters 0‑5 in an Aspose.Cells textbox. | Show how to apply distinct font attributes to a substring of a textbox and then set the paragraph's AlignmentType to Center with Aspose.Cells. | Explain the steps to retrieve TextParagraph objects from a shape's TextBody and modify their alignment for partial text in an Excel file.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

namespace AsposeCellsExamples
{
    // Creates a workbook, adds a TextBox shape, applies a blue Arial style to the first five characters, centers the paragraph that contains those characters, and saves the file as an Excel workbook using Aspose.Cells for .NET.
    public class CenterAlignSelectedCharactersInTextBox
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
            Shape textBox = worksheet.Shapes.AddTextBox(0, 0, 100, 100, 300, 100);
            // Set the text of the text box
            textBox.Text = "Hello World";

            // Get the FontSettingCollection (represents the rich‑text content of the shape)
            FontSettingCollection fontSettings = textBox.TextBody;

            // Define a style for characters 0‑5 ("Hello")
            Style style = workbook.CreateStyle();
            style.Font.Name = "Arial";
            style.Font.Size = 14;
            style.Font.Color = System.Drawing.Color.Blue;

            StyleFlag flag = new StyleFlag
            {
                FontName = true,
                FontSize = true,
                FontColor = true
            };

            // Apply the style to characters 0‑5
            fontSettings.Format(0, 5, style.Font, flag);

            // Retrieve the first paragraph and center‑align it
            TextParagraph paragraph = fontSettings.TextParagraphs[0];
            paragraph.AlignmentType = TextAlignmentType.Center;

            // Save the workbook
            string outputPath = "CenterAlignedTextBox.xlsx";
            workbook.Save(outputPath);

            Console.WriteLine($"Workbook saved with centered characters inside the text box: {outputPath}");
        }
    }
}
