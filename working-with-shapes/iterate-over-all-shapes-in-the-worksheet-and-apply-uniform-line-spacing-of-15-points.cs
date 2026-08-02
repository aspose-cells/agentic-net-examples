// Title: Set 1.5‑point line spacing for all rich‑text shapes in Aspose.Cells (C#)
// Description: Creates a workbook, adds a multi‑line text box and a rectangle, then loops through every shape on the first worksheet. For each rich‑text shape it accesses the TextBody, iterates its paragraphs, switches the spacing unit to points and applies a 1.5‑point line gap before saving the file.
// Keywords: Aspose.Cells line spacing C# | shape text paragraph spacing | rich text shape Aspose.Cells | uniform line spacing Excel shape | modify text box line height .NET
// Common Searches: C# Aspose.Cells set line spacing for shape text | How to change paragraph spacing in Excel text box using Aspose | Iterate worksheet shapes and adjust line spacing | Apply consistent line height to all rich‑text shapes in a workbook | Aspose.Cells example for uniform shape paragraph spacing
// Developer Intent: Apply a consistent 1.5‑point line spacing to every paragraph inside each rich‑text shape on a worksheet.
// Use Cases: Standardize appearance of multi‑line text boxes in auto‑generated reports. | Ensure uniform paragraph spacing before exporting to PDF or image formats. | Enforce corporate style rules that require a specific line‑spacing value for shape text.
// AI Prompts: Generate C# code that sets a variable line‑spacing value for all paragraphs in rich‑text shapes with Aspose.Cells. | Show how to revert line spacing to the default for shapes that do not contain rich text. | Provide an example that changes line spacing to 2 points only for text boxes, leaving other shapes untouched.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

// Creates a workbook, adds a multi‑line text box and a rectangle, then loops through every shape on the first worksheet. For each rich‑text shape it accesses the TextBody, iterates its paragraphs, switches the spacing unit to points and applies a 1.5‑point line gap before saving the file.
class UniformLineSpacingDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add sample shapes (a text box with multiple lines and a rectangle)
        Shape textBox = worksheet.Shapes.AddTextBox(0, 0, 0, 0, 300, 100);
        textBox.Text = "First line\nSecond line\nThird line";

        worksheet.Shapes.AddRectangle(2, 0, 2, 0, 150, 100); // non‑text shape

        // Iterate over all shapes in the worksheet
        foreach (Shape shape in worksheet.Shapes)
        {
            // Process only shapes that contain rich text (e.g., text boxes)
            if (shape.IsRichText && shape.TextBody != null)
            {
                // Access the collection of paragraphs within the shape's text body
                TextParagraphCollection paragraphs = shape.TextBody.TextParagraphs;

                // Apply uniform line spacing of 1.5 points to each paragraph
                foreach (TextParagraph paragraph in paragraphs)
                {
                    paragraph.LineSpaceSizeType = LineSpaceSizeType.Points; // use points as unit
                    paragraph.LineSpace = 1.5; // set line spacing to 1.5 points
                }
            }
        }

        // Save the workbook with the updated line spacing
        workbook.Save("UniformLineSpacing.xlsx");
    }
}
