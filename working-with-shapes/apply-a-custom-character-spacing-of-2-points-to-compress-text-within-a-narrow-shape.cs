// Title: Compress Text with Negative Character Spacing in a Narrow TextBox using Aspose.Cells for .NET
// Description: Creates a workbook, adds a 50 × 100 pt textbox shape, sets TextOptions.Spacing to -2 points to compress the text, optionally calls FitToTextSize to resize the shape, and saves the file as NegativeSpacingDemo.xlsx.
// Keywords: Aspose.Cells | .NET | negative character spacing | text box shape | TextOptions.Spacing | FitToTextSize | compress text | narrow shape
// Common Searches: Aspose.Cells set negative spacing for shape text | compress text in narrow textbox Aspose.Cells .NET | Fit shape to text after spacing adjustment Aspose.Cells | How to use TextOptions.Spacing in Aspose.Cells
// Developer Intent: Compress the text inside a narrow textbox shape by applying a negative character spacing value.
// Use Cases: Tight label creation for compact reports | Space‑saving legends or annotations in charts | Fitting headings into small cells of printable forms
// AI Prompts: Demonstrate setting TextOptions.Spacing to a negative value and calling FitToTextSize for a textbox shape in Aspose.Cells .NET. | Show how to apply different negative spacing values to multiple shapes in a worksheet and save the workbook. | Explain how to retrieve the current spacing, adjust it conditionally, and update the shape dimensions accordingly.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

// Creates a workbook, adds a 50 × 100 pt textbox shape, sets TextOptions.Spacing to -2 points to compress the text, optionally calls FitToTextSize to resize the shape, and saves the file as NegativeSpacingDemo.xlsx.
class ApplyNegativeCharacterSpacing
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a narrow text box shape (width = 50 points, height = 100 points)
        Shape textBox = worksheet.Shapes.AddTextBox(0, 0, 0, 0, 50, 100);
        textBox.Text = "Compressed Text Example";

        // Access the TextOptions of the shape
        TextOptions textOptions = textBox.TextOptions;

        // Apply custom character spacing of -2 points to compress the text
        textOptions.Spacing = -2.0;

        // Recalculate the shape size to fit the adjusted text (optional)
        textBox.FitToTextSize();

        // Save the workbook
        workbook.Save("NegativeSpacingDemo.xlsx");
    }
}
