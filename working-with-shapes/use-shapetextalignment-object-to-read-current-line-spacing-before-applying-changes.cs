// Title: C# – Read a shape's paragraph line spacing before modifying ShapeTextAlignment with Aspose.Cells
// Description: Demonstrates how to retrieve the current line‑spacing value of a text box paragraph, then adjust wrapping and margin settings via the ShapeTextAlignment object, save the workbook, and confirm that the original spacing is retained after reload.
// Keywords: Aspose.Cells C# shape line spacing | ShapeTextAlignment margins | read paragraph line space Aspose.Cells | preserve line spacing shape | text box formatting .NET
// Common Searches: Aspose.Cells get line spacing of shape text | C# read paragraph line space before ShapeTextAlignment | How to check line spacing of a text box shape in Aspose.Cells | Retrieve and keep shape line spacing after formatting
// Developer Intent: Obtain the existing line‑spacing value of a shape's text paragraph before applying ShapeTextAlignment properties.
// Use Cases: Log original line spacing before applying alignment changes. | Copy shapes between worksheets while maintaining paragraph spacing. | Validate that margin adjustments do not alter line spacing after saving.
// AI Prompts: Generate C# code that reads the line spacing of the first paragraph in a text box shape, then sets ShapeTextAlignment margins and enables wrapping using Aspose.Cells. | Show how to compare the paragraph line spacing before and after saving a workbook that contains a formatted shape. | Explain best practices for preserving paragraph line spacing when programmatically updating shape alignment in Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

// Demonstrates how to retrieve the current line‑spacing value of a text box paragraph, then adjust wrapping and margin settings via the ShapeTextAlignment object, save the workbook, and confirm that the original spacing is retained after reload.
class Program
{
    static void Main()
    {
        // Create a new workbook (creation rule)
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add a text box shape to the worksheet
        Shape shape = sheet.Shapes.AddTextBox(0, 0, 0, 0, 300, 100);
        shape.Text = "First line\nSecond line";

        // Access the first paragraph of the shape's text body
        TextParagraph paragraph = shape.TextBody.TextParagraphs[0];
        // Ensure the line space unit is points
        paragraph.LineSpaceSizeType = LineSpaceSizeType.Points;

        // Read the current line spacing value before making any changes
        double currentLineSpace = paragraph.LineSpace;
        Console.WriteLine("Current line spacing (points): " + currentLineSpace);

        // Modify shape's text alignment using the ShapeTextAlignment object
        ShapeTextAlignment textAlignment = shape.TextBody.TextAlignment;
        textAlignment.IsTextWrapped = true;          // enable text wrapping
        textAlignment.TopMarginPt = 12;              // set top margin in points
        textAlignment.BottomMarginPt = 12;           // set bottom margin in points
        textAlignment.LeftMarginPt = 6;              // set left margin in points
        textAlignment.RightMarginPt = 6;             // set right margin in points

        // Save the workbook (save rule)
        string filePath = "ShapeTextAlignmentLineSpaceDemo.xlsx";
        workbook.Save(filePath);

        // Load the workbook back (load rule) to verify the saved settings
        Workbook loadedWorkbook = new Workbook(filePath);
        Shape loadedShape = loadedWorkbook.Worksheets[0].Shapes[0];
        double loadedLineSpace = loadedShape.TextBody.TextParagraphs[0].LineSpace;
        Console.WriteLine("Loaded line spacing (points): " + loadedLineSpace);
    }
}
