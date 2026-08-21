// Title: C# – Set 1.5 line spacing in an Aspose.Cells textbox shape using TextParagraph and TextOptions
// Description: Creates a workbook, adds a textbox shape, inserts multiline text, and applies 1.5‑line (150 %) spacing via TextParagraph.LineSpaceSizeType = Percentage and LineSpace = 150 (or TextOptions.LineSpacing). The shape is then auto‑sized with FitToTextSize and saved as an XLSX file.
// Keywords: Aspose.Cells line spacing | textbox shape line spacing .NET | TextParagraph LineSpacePercentage | TextOptions.LineSpacing Aspose | C# adjust shape text layout | fit shape to text Aspose.Cells | 1.5 line spacing Excel shape | Aspose.Cells shape formatting
// Common Searches: how to change line spacing in a textbox shape using Aspose.Cells C# | Aspose.Cells TextParagraph line spacing percentage example | set 150% line spacing for shape text in .NET | auto resize shape after modifying line spacing Aspose | TextOptions.LineSpacing usage in Aspose.Cells
// Developer Intent: Apply 1.5‑line (150 %) spacing to the text inside a shape.
// Use Cases: Design a readable multiline textbox in an Excel worksheet. | Maintain consistent paragraph spacing across multiple shapes. | Resize a shape automatically after changing line spacing to avoid clipping.
// AI Prompts: Generate C# code that sets 150% line spacing for every paragraph in a shape's TextBody using Aspose.Cells. | Show how to loop through all TextParagraphs of a shape and apply 1.5 line spacing, then call FitToTextSize. | Create an example that reads a line‑spacing value from a variable and applies it via TextOptions.LineSpacing to a textbox shape.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

// Creates a workbook, adds a textbox shape, inserts multiline text, and applies 1.5‑line (150 %) spacing via TextParagraph.LineSpaceSizeType = Percentage and LineSpace = 150 (or TextOptions.LineSpacing). The shape is then auto‑sized with FitToTextSize and saved as an XLSX file.
class AdjustLineSpacingDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a text box shape to the worksheet
        // Parameters: upper left row, upper left column, upper left offset in pixels,
        // lower right row, lower right column, lower right offset in pixels
        Shape textBox = worksheet.Shapes.AddTextBox(2, 2, 0, 2, 2, 0);
        textBox.Width = 300;
        textBox.Height = 150;

        // Set multiline text inside the shape
        textBox.Text = "First line\nSecond line\nThird line";

        // Access the first paragraph of the text box
        TextParagraph paragraph = textBox.TextBody.TextParagraphs[0];

        // Set line spacing to 1.5 lines (150% of the font size)
        paragraph.LineSpaceSizeType = LineSpaceSizeType.Percentage; // use percentage unit
        paragraph.LineSpace = 150; // 150% = 1.5 lines

        // Optionally, fit the shape size to the new text layout
        textBox.FitToTextSize();

        // Save the workbook
        workbook.Save("AdjustedLineSpacing.xlsx");
    }
}
