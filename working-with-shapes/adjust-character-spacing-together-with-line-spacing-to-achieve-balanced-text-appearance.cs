// Title: Adjust character and line spacing in a textbox shape using Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook, add a textbox shape, set multi‑line text, increase character spacing via TextOptions.Spacing, and apply precise line spacing (points) to each paragraph, resulting in a balanced visual layout inside the shape.
// Keywords: Aspose.Cells C# | textbox shape spacing | character spacing Aspose.Cells | line spacing Excel shape | TextOptions.Spacing | TextParagraph.LineSpace | Excel shape text formatting | .NET spreadsheet API | adjust text appearance
// Common Searches: how to set character spacing in an Excel textbox with Aspose.Cells | line spacing for text paragraphs in a shape using Aspose.Cells .NET | increase letter spacing and line height in a shape programmatically | Aspose.Cells example for text formatting inside shapes | C# code to adjust spacing of textbox content in Excel
// Developer Intent: Modify both the inter‑character distance and the inter‑line distance of text inside a textbox shape to achieve a visually balanced appearance.
// Use Cases: Designing title boxes with enhanced readability for dashboards. | Generating multi‑line comments or notes in reports where uniform spacing is required. | Creating certificates or awards where text inside shapes must follow strict typographic standards.
// AI Prompts: Show C# code to set TextOptions.Spacing to 2.0 and paragraph line spacing to 8 points for all paragraphs in a textbox shape. | Provide a reusable method that accepts character‑spacing and line‑spacing parameters and applies them to a shape's text. | Explain how to read the current line‑spacing settings of a textbox shape and modify them with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

// Demonstrates how to create a workbook, add a textbox shape, set multi‑line text, increase character spacing via TextOptions.Spacing, and apply precise line spacing (points) to each paragraph, resulting in a balanced visual layout inside the shape.
class AdjustSpacingDemo
{
    public static void Main()
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
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a text box shape to the worksheet
        // Parameters: upper left row, upper left column, top offset, left offset, width, height
        Shape textBox = worksheet.Shapes.AddTextBox(0, 0, 0, 0, 300, 150);
        textBox.Text = "First line\nSecond line\nThird line";

        // Adjust character spacing for the entire text run
        TextOptions textOptions = textBox.TextOptions;
        textOptions.Spacing = 1.5; // increase spacing between characters

        // Access all paragraphs inside the text box
        TextParagraphCollection paragraphs = textBox.TextBody.TextParagraphs;

        // Apply line spacing to each paragraph to balance appearance
        foreach (TextParagraph paragraph in paragraphs)
        {
            paragraph.LineSpaceSizeType = LineSpaceSizeType.Points; // use points as unit
            paragraph.LineSpace = 5; // set line spacing to 5 points
        }

        // Save the workbook with the adjusted spacing
        workbook.Save("AdjustedSpacingDemo.xlsx");
    }
}
