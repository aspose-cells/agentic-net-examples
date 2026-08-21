// Title: C# utility to apply paragraph formatting to an Aspose.Cells Shape TextBody
// Description: A static method that validates a Shape, iterates its TextParagraph collection and sets center alignment, 18‑point line spacing, first‑line indent of 10, left/right margins of 5, and space before/after of 4, then calls FitToTextSize to resize the shape.
// Keywords: Aspose.Cells shape formatting | C# paragraph formatting shape | text alignment Aspose.Cells | line spacing shape Aspose | shape margins C# | FitToTextSize Aspose.Cells
// Common Searches: how to set paragraph alignment for a shape in Aspose.Cells C# | apply line spacing and margins to shape text Aspose.Cells | C# utility to format shape TextBody paragraphs | resize shape to fit formatted text Aspose.Cells | Aspose.Cells shape text formatting example
// Developer Intent: Apply consistent paragraph formatting to all text paragraphs of a Shape and automatically resize the shape to fit the formatted content.
// Use Cases: Standardize callout or text‑box appearance across generated reports. | Prepare template workbooks with uniform paragraph settings before populating data. | Prevent text clipping by resizing shapes after applying paragraph styles.
// AI Prompts: Write a C# method that receives an Aspose.Cells Shape and sets its TextParagraph properties for alignment, line spacing, indents, and margins, then calls FitToTextSize. | Show how to loop through all shapes on a worksheet and apply the ApplyParagraphFormatting utility to each one. | Explain how to refactor the utility so formatting values are passed as parameters instead of being hard‑coded.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

// A static method that validates a Shape, iterates its TextParagraph collection and sets center alignment, 18‑point line spacing, first‑line indent of 10, left/right margins of 5, and space before/after of 4, then calls FitToTextSize to resize the shape.
public static class ShapeFormattingUtility
{
    /// <param name="shape">The shape whose text paragraphs will be formatted.</param>
    public static void ApplyParagraphFormatting(Shape shape)
    {
        if (shape == null) return;
        if (shape.TextBody == null || shape.TextBody.TextParagraphs == null)
            return;

        const TextAlignmentType alignment = TextAlignmentType.Center;
        const LineSpaceSizeType lineSpaceType = LineSpaceSizeType.Points;
        const double lineSpace = 18;
        const double firstLineIndent = 10;
        const double leftMargin = 5;
        const double rightMargin = 5;
        const double spaceBefore = 4;
        const double spaceAfter = 4;

        foreach (TextParagraph paragraph in shape.TextBody.TextParagraphs)
        {
            paragraph.AlignmentType = alignment;
            paragraph.LineSpaceSizeType = lineSpaceType;
            paragraph.LineSpace = lineSpace;
            paragraph.FirstLineIndent = firstLineIndent;
            paragraph.LeftMargin = leftMargin;
            paragraph.RightMargin = rightMargin;
            paragraph.SpaceBefore = spaceBefore;
            paragraph.SpaceAfter = spaceAfter;
        }

        shape.FitToTextSize();
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        try
        {
            string inputPath = "SampleWorkbook.xlsx";

            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                // Create a new workbook with a shape for demonstration
                workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Shape shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 1, 1, 0, 0, 200, 100);
                shape.Text = "First line.\nSecond line.";
            }

            Worksheet ws = workbook.Worksheets[0];
            if (ws.Shapes.Count > 0)
            {
                Shape targetShape = ws.Shapes[0];
                ShapeFormattingUtility.ApplyParagraphFormatting(targetShape);
            }

            string outputPath = "FormattedWorkbook.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
