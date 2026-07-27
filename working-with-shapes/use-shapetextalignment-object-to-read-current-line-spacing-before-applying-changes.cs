// Title: C# – Read and Adjust TextBox Line Spacing and Margins with Aspose.Cells ShapeTextAlignment
// Description: Demonstrates how to retrieve the current LineSpace and LineSpaceSizeType of a textbox shape's first TextParagraph, change the line spacing, modify top/bottom margins via ShapeTextAlignment, and verify the settings after saving the workbook.
// Keywords: Aspose.Cells C# | ShapeTextAlignment | TextParagraph LineSpace | textbox line spacing | modify shape margins | .NET spreadsheet shape formatting | read line spacing Aspose.Cells | set TopMarginPt BottomMarginPt
// Common Searches: Aspose.Cells read textbox line spacing C# | change shape text margins with ShapeTextAlignment | how to get LineSpaceSizeType in Aspose.Cells | modify multiline textbox line spacing .NET | persist shape text formatting after save Aspose.Cells
// Developer Intent: Retrieve a shape's existing line spacing, then apply new line spacing and margin values using ShapeTextAlignment.
// Use Cases: Inspect the current LineSpace and its size type before applying formatting changes. | Set a specific point value for line spacing of multiline text inside a textbox shape. | Adjust top and bottom margins of a shape's text body via ShapeTextAlignment and confirm persistence after workbook save/load.
// AI Prompts: Generate C# code that reads the LineSpace and LineSpaceSizeType of a textbox shape's first TextParagraph, then sets LineSpace to 12 points and updates ShapeTextAlignment.TopMarginPt and BottomMarginPt to 15 points. | Provide a snippet that creates a textbox shape, prints its default line spacing, changes the spacing to 8 points, adjusts top/bottom margins, saves the workbook, reloads it, and outputs the verified values. | Explain how TextParagraph line‑spacing properties interact with ShapeTextAlignment margins in Aspose.Cells and how to ensure those changes are saved in the workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

// Demonstrates how to retrieve the current LineSpace and LineSpaceSizeType of a textbox shape's first TextParagraph, change the line spacing, modify top/bottom margins via ShapeTextAlignment, and verify the settings after saving the workbook.
class ShapeTextAlignmentLineSpacingDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a text box shape with some multiline text
        Shape shape = worksheet.Shapes.AddTextBox(0, 0, 0, 0, 300, 200);
        shape.Text = "First line\nSecond line";

        // Access the first paragraph of the shape's text body
        TextParagraph paragraph = shape.TextBody.TextParagraphs[0];

        // Read and display the current line spacing and its size type
        double currentLineSpace = paragraph.LineSpace;
        LineSpaceSizeType currentSizeType = paragraph.LineSpaceSizeType;
        Console.WriteLine($"Current LineSpace: {currentLineSpace}, SizeType: {currentSizeType}");

        // Apply new line spacing settings
        paragraph.LineSpaceSizeType = LineSpaceSizeType.Points;
        paragraph.LineSpace = 10; // Set line spacing to 10 points

        // Access the ShapeTextAlignment object to read and modify margin settings
        ShapeTextAlignment textAlignment = shape.TextBody.TextAlignment;
        Console.WriteLine($"Default TopMarginPt: {textAlignment.TopMarginPt}");
        textAlignment.TopMarginPt = 20;
        textAlignment.BottomMarginPt = 20;

        // Save the workbook
        string outputPath = "ShapeTextAlignmentLineSpacingDemo.xlsx";
        workbook.Save(outputPath);

        // Load the workbook to verify the applied settings
        Workbook loadedWorkbook = new Workbook(outputPath);
        Shape loadedShape = loadedWorkbook.Worksheets[0].Shapes[0];
        TextParagraph loadedParagraph = loadedShape.TextBody.TextParagraphs[0];
        Console.WriteLine($"Loaded LineSpace: {loadedParagraph.LineSpace}");
        ShapeTextAlignment loadedAlignment = loadedShape.TextBody.TextAlignment;
        Console.WriteLine($"Loaded TopMarginPt: {loadedAlignment.TopMarginPt}");
    }
}
