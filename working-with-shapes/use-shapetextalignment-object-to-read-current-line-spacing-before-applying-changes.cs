using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

class ShapeTextAlignmentLineSpaceDemo
{
    public static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a text box shape to the worksheet
        Shape shape = worksheet.Shapes.AddTextBox(0, 0, 0, 0, 300, 100);
        shape.Text = "First line\nSecond line";

        // Obtain the ShapeTextAlignment object (required by the task)
        ShapeTextAlignment textAlignment = shape.TextBody.TextAlignment;

        // Access the first paragraph of the shape's text body
        TextParagraph paragraph = shape.TextBody.TextParagraphs[0];

        // Read and display the current line spacing before any changes
        double currentLineSpace = paragraph.LineSpace;
        LineSpaceSizeType currentSizeType = paragraph.LineSpaceSizeType;
        Console.WriteLine($"Current LineSpace: {currentLineSpace}");
        Console.WriteLine($"Current LineSpaceSizeType: {currentSizeType}");

        // Apply new line spacing values
        paragraph.LineSpaceSizeType = LineSpaceSizeType.Points;
        paragraph.LineSpace = 12; // Set line spacing to 12 points

        // Save the workbook to a file
        string outputPath = "ShapeTextAlignmentLineSpaceDemo.xlsx";
        workbook.Save(outputPath);
        Console.WriteLine($"Workbook saved to {outputPath}");
    }
}