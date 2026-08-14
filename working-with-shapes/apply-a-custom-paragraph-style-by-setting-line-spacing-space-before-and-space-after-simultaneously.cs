// Title: C# – Set Line Spacing, Space Before & After for a TextBox Paragraph with Aspose.Cells
// Description: Demonstrates how to create a workbook, add a textbox shape with two paragraphs, and apply custom line spacing (points), space before, and space after to the second paragraph. The example saves the file, reloads it, and prints the paragraph settings to confirm persistence.
// Keywords: Aspose.Cells C# paragraph spacing | textbox line spacing Aspose.Cells | SpaceBefore Aspose.Cells | SpaceAfter Aspose.Cells | LineSpaceSizeType Points | Excel shape text formatting | Aspose.Cells paragraph style example | C# Excel textbox paragraph
// Common Searches: how to set line spacing in a textbox paragraph using Aspose.Cells C# | Aspose.Cells set SpaceBefore and SpaceAfter for a paragraph | C# Aspose.Cells custom paragraph formatting in shapes | apply paragraph spacing to Excel textbox with Aspose.Cells | verify paragraph style persistence after saving Aspose.Cells workbook
// Developer Intent: Apply custom line spacing, space before, and space after to a specific paragraph inside a textbox shape in an Excel workbook using Aspose.Cells for .NET.
// Use Cases: Design reports where each paragraph in a textbox requires distinct spacing for visual clarity. | Generate Excel templates that match printed document layouts with precise paragraph spacing. | Programmatically ensure paragraph formatting survives workbook save and reload cycles.
// AI Prompts: Show how to use other line‑spacing units (percent, auto) for a textbox paragraph in Aspose.Cells. | Provide code to loop through all paragraphs in a textbox and set identical spacing values. | Explain how to copy paragraph formatting from one shape to another using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

namespace AsposeCellsParagraphStyleDemo
{
    // Demonstrates how to create a workbook, add a textbox shape with two paragraphs, and apply custom line spacing (points), space before, and space after to the second paragraph. The example saves the file, reloads it, and prints the paragraph settings to confirm persistence.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook (creation rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a text box shape to the worksheet
            // Parameters: upper left row, upper left column, top, left, width, height
            Shape textBox = worksheet.Shapes.AddTextBox(0, 0, 0, 0, 400, 200);

            // Set multi‑paragraph text (two paragraphs separated by newline)
            textBox.Text = "First paragraph\nSecond paragraph";

            // Access the collection of paragraphs inside the text box
            TextParagraphCollection paragraphs = textBox.TextBody.TextParagraphs;

            // Choose the second paragraph (index 1) to apply custom spacing
            TextParagraph paragraph = paragraphs[1];

            // Set line spacing unit to points and define the line spacing value
            paragraph.LineSpaceSizeType = LineSpaceSizeType.Points;
            paragraph.LineSpace = 12; // 12 points line spacing

            // Set space before and after the paragraph (also in points)
            paragraph.SpaceBefore = 8;   // 8 points before the paragraph
            paragraph.SpaceAfter = 10;   // 10 points after the paragraph

            // Save the workbook (save rule)
            string outputPath = "ParagraphStyleDemo.xlsx";
            workbook.Save(outputPath);

            // Optional: Load the workbook again to verify that settings persist (load rule)
            Workbook loadedWorkbook = new Workbook(outputPath);
            Shape loadedShape = loadedWorkbook.Worksheets[0].Shapes[0];
            TextParagraph loadedParagraph = loadedShape.TextBody.TextParagraphs[1];

            Console.WriteLine("Loaded Paragraph Settings:");
            Console.WriteLine($"LineSpaceSizeType: {loadedParagraph.LineSpaceSizeType}");
            Console.WriteLine($"LineSpace: {loadedParagraph.LineSpace}");
            Console.WriteLine($"SpaceBefore: {loadedParagraph.SpaceBefore}");
            Console.WriteLine($"SpaceAfter: {loadedParagraph.SpaceAfter}");
        }
    }
}
