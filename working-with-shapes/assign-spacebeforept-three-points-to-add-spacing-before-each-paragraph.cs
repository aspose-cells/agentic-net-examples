// Title: Add 3‑point spacing before each paragraph in a TextBox with Aspose.Cells for .NET
// Description: Creates a workbook, inserts a TextBox shape, fills it with three newline‑separated paragraphs, and sets the SpaceBefore property to 3 points for every paragraph before saving the file.
// Keywords: Aspose.Cells C# text box spacing | SpaceBefore property Aspose.Cells | LineSpaceSizeType Points | Excel shape paragraph formatting | C# set paragraph spacing | Aspose.Cells .NET API | global developers
// Common Searches: Aspose.Cells set SpaceBefore C# | Add space before paragraph in Excel shape | TextBox paragraph spacing Aspose.Cells | Change line spacing in Aspose.Cells text box | C# example paragraph spacing Excel
// Developer Intent: Apply a 3‑point SpaceBefore value to every paragraph inside a TextBox shape in an Excel workbook using Aspose.Cells for .NET.
// Use Cases: Generate reports where notes in a TextBox need consistent leading for readability. | Create templates with multi‑line instructions, ensuring uniform space before each paragraph. | Match corporate style guidelines by programmatically adjusting paragraph spacing in shape‑based comments.
// AI Prompts: Write C# code with Aspose.Cells to set SpaceBefore to 5 points for all paragraphs in a TextBox shape. | Explain how to switch the line‑spacing unit from points to centimeters for TextParagraph objects in Aspose.Cells. | Show how to access a specific paragraph by index in a TextBox and modify its SpaceBefore value.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

namespace AsposeCellsExamples
{
    // Creates a workbook, inserts a TextBox shape, fills it with three newline‑separated paragraphs, and sets the SpaceBefore property to 3 points for every paragraph before saving the file.
    public class SetSpaceBeforeDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a text box shape to the worksheet
                // Parameters: upper left row, upper left column, upper left row offset, upper left column offset, width, height
                Shape textBox = worksheet.Shapes.AddTextBox(0, 0, 0, 0, 300, 200);
                // Set text with multiple paragraphs (separated by newline)
                textBox.Text = "First paragraph\nSecond paragraph\nThird paragraph";

                // Access the collection of paragraphs in the text box
                TextParagraphCollection paragraphs = textBox.TextBody.TextParagraphs;

                // Apply a spacing of 3 points before each paragraph
                foreach (TextParagraph paragraph in paragraphs)
                {
                    // Ensure the spacing unit is points (optional, default is points)
                    paragraph.SpaceBeforeSizeType = LineSpaceSizeType.Points;
                    // Set the amount of space before the paragraph
                    paragraph.SpaceBefore = 3.0;
                }

                // Save the workbook to a file
                workbook.Save("SetSpaceBeforeDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            SetSpaceBeforeDemo.Run();
        }
    }
}
