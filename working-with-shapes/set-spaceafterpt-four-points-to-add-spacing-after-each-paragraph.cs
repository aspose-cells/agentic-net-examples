// Title: Aspose.Cells for .NET – Set a 4‑point SpaceAfter on every paragraph in a TextBox shape
// Description: This example creates a workbook, inserts a TextBox shape with multiline text, accesses its TextParagraphCollection, and sets each paragraph's SpaceAfterSizeType to Points and SpaceAfter to 4. The workbook is then saved as SetSpaceAfterDemo.xlsx.
// Keywords: Aspose.Cells .NET paragraph spacing | SpaceAfter property | TextBox shape line spacing | LineSpaceSizeType Points | Excel shape paragraph formatting | C# Aspose.Cells example
// Common Searches: how to add space after paragraphs in Aspose.Cells | set paragraph spacing in a text box using Aspose.Cells .NET | Aspose.Cells SpaceAfterSizeType example | C# code to adjust line spacing in Excel shapes | configure paragraph spacing in Aspose.Cells workbook
// Developer Intent: Apply a uniform 4‑point spacing after each paragraph inside a TextBox shape in an Excel file using Aspose.Cells for .NET.
// Use Cases: Design reports where text boxes need consistent paragraph spacing for readability. | Create reusable Excel templates with pre‑formatted text boxes that match corporate style guides. | Automate generation of documentation worksheets that require exact spacing after each paragraph.
// AI Prompts: Show me how to set SpaceAfter to 6 points for all paragraphs in a TextBox with Aspose.Cells .NET. | Provide C# code to change SpaceAfterSizeType to Auto and calculate spacing based on paragraph length. | Explain how to read and modify existing paragraph spacing in a saved Excel workbook using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

namespace AsposeCellsExamples
{
    // This example creates a workbook, inserts a TextBox shape with multiline text, accesses its TextParagraphCollection, and sets each paragraph's SpaceAfterSizeType to Points and SpaceAfter to 4. The workbook is then saved as SetSpaceAfterDemo.xlsx.
    public class SetSpaceAfterDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a text box shape to the worksheet
                Shape textBox = worksheet.Shapes.AddTextBox(0, 0, 0, 0, 300, 200);
                // Set text with multiple paragraphs
                textBox.Text = "First paragraph\nSecond paragraph\nThird paragraph";

                // Access the collection of paragraphs in the text box
                TextParagraphCollection paragraphs = textBox.TextBody.TextParagraphs;

                // Apply 4 points spacing after each paragraph
                foreach (TextParagraph paragraph in paragraphs)
                {
                    // Use points as the unit for space after
                    paragraph.SpaceAfterSizeType = LineSpaceSizeType.Points;
                    // Set the space after value to 4 points
                    paragraph.SpaceAfter = 4;
                }

                // Save the workbook to a file
                workbook.Save("SetSpaceAfterDemo.xlsx");
                Console.WriteLine("Workbook saved successfully as SetSpaceAfterDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            SetSpaceAfterDemo.Run();
        }
    }
}
