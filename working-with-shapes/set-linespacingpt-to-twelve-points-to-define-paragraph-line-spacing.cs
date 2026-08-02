// Title: Aspose.Cells C# – Set Text Box Paragraph Line Spacing to 12 Points
// Description: Creates a workbook, inserts a text box, accesses its first paragraph, sets the line‑spacing unit to points, applies a 12‑point spacing, and saves the file as an Excel workbook.
// Keywords: Aspose.Cells line spacing | C# text box paragraph spacing | LineSpaceSizeType Points | Excel shape text formatting | .NET set paragraph line height | Aspose.Cells TextParagraph example
// Common Searches: Aspose.Cells set paragraph line spacing points | C# change line height in Excel text box | How to use LineSpaceSizeType with Aspose.Cells | Set 12 point line spacing in shape text | Aspose.Cells text box formatting tutorial
// Developer Intent: Apply a 12‑point line spacing to a text box paragraph in an Excel sheet using Aspose.Cells for .NET.
// Use Cases: Generate reports where text boxes follow a strict 12‑point line‑spacing rule for brand consistency. | Improve readability of multi‑line notes inside shapes on dashboards created programmatically. | Automate creation of Excel templates that require precise typographic layout for corporate documentation.
// AI Prompts: Give C# code that changes the line spacing of every paragraph in a shape to 10 points with Aspose.Cells. | Show how to set the line‑spacing unit to points and also align paragraphs inside a text box using Aspose.Cells for .NET. | Explain how to assign different line‑spacing values to multiple paragraphs within the same text box shape.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

// Creates a workbook, inserts a text box, accesses its first paragraph, sets the line‑spacing unit to points, applies a 12‑point spacing, and saves the file as an Excel workbook.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a text box shape to the worksheet
        Shape textBox = worksheet.Shapes.AddTextBox(0, 0, 0, 0, 400, 200);
        textBox.Text = "First line\nSecond line";

        // Access the first paragraph of the text box
        TextParagraph paragraph = textBox.TextBody.TextParagraphs[0];

        // Define line spacing in points
        paragraph.LineSpaceSizeType = LineSpaceSizeType.Points; // Use points as the unit
        paragraph.LineSpace = 12; // Set line spacing to twelve points

        // Save the workbook
        workbook.Save("LineSpacingDemo.xlsx");
    }
}
